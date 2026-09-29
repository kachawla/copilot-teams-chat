import { createHash } from "node:crypto";
import { openPrepared, prepareChat } from "./compose.mjs";
import { startComposeServer } from "./compose-server.mjs";
import { DraftStore, applyDraftPatch, validDraftId } from "./drafts.mjs";
import { describeError } from "./service.mjs";
import { COMPOSE_CANVAS_ID } from "./tools.mjs";

const draftProperties = {
    to: { type: "array", items: { type: "string" }, description: "People to message: email addresses or saved contact names." },
    topic: { type: "string", description: "Name for a new group chat." },
    message: { type: "string", description: "Draft message (Markdown unless format is text)." },
    format: { type: "string", enum: ["markdown", "text"] },
};

/**
 * Compose canvas: the agent fills in a draft, the user edits it, then clicks
 * "Open in Teams" and presses Send there. Nothing is sent from the canvas.
 */
export function createComposeCanvas({ createCanvas, CanvasError, service, session, draftsDir }) {
    let sessionId;
    const drafts = new DraftStore(() => draftsDir(sessionId));
    const instances = new Map(); // instanceId -> { server, draftId }

    const broadcastAll = async (draftId) => {
        for (const entry of instances.values()) {
            if (!draftId || entry.draftId === draftId) await entry.server.broadcast().catch(() => {});
        }
    };

    const stateFor = async (draftId) => ({
        draft: await drafts.get(draftId),
        contacts: (await service.contacts.list()).map((c) => c.name),
    });

    const updateDraft = async (draftId, patch, extra = {}) => {
        const draft = { ...applyDraftPatch(await drafts.get(draftId), patch), ...extra };
        await drafts.save(draft);
        await broadcastAll(draftId);
        return drafts.get(draftId);
    };

    const handleOpen = async (draftId) => {
        const draft = await drafts.get(draftId);
        try {
            const prepared = await prepareChat({ service, to: draft.to, topic: draft.topic, message: draft.message, format: draft.format });
            const openedIn = await openPrepared({ service, prepared });
            await drafts.save({
                ...draft,
                lastResult: { openedIn, label: prepared.label, link: prepared.webLink, at: new Date().toISOString() },
            });
            await broadcastAll(draftId);
            if (openedIn) await session.log?.(`Opened a Teams draft to ${prepared.label} from the compose canvas.`);
            return { openedIn, label: prepared.label, link: prepared.webLink, state: await stateFor(draftId) };
        } catch (err) {
            if (err.code === "unknown_recipient") {
                throw new Error(
                    `No email address is saved for ${err.names.map((n) => `"${n}"`).join(", ")}. Use their email address, or ask Copilot to save them as a contact.`,
                );
            }
            throw new Error(describeError(err));
        }
    };

    const draftIdFor = (instanceId) => instances.get(instanceId)?.draftId ?? "default";

    return createCanvas({
        id: COMPOSE_CANVAS_ID,
        displayName: "Teams message",
        description: "Draft a Microsoft Teams message with the user, then open it in Teams where they press Send.",
        inputSchema: {
            type: "object",
            properties: {
                draftId: { type: "string", description: 'Draft to open. Defaults to "default". Use different IDs for separate drafts.' },
                ...draftProperties,
            },
        },
        actions: [
            {
                name: "update_draft",
                description: "Replace fields of the draft shown in this panel (to, topic, message, format).",
                inputSchema: { type: "object", properties: draftProperties },
                handler: async (ctx) => {
                    sessionId ??= ctx.sessionId;
                    return updateDraft(draftIdFor(ctx.instanceId), ctx.input ?? {});
                },
            },
            {
                name: "get_draft",
                description: "Read the current draft, including the user's edits.",
                handler: async (ctx) => {
                    sessionId ??= ctx.sessionId;
                    return drafts.get(draftIdFor(ctx.instanceId));
                },
            },
        ],
        open: async (ctx) => {
            sessionId ??= ctx.sessionId;
            const input = ctx.input ?? {};
            const draftId = input.draftId ?? "default";
            if (!validDraftId(draftId)) throw new CanvasError("invalid_draft_id", `Invalid draftId "${draftId}".`);

            // The runtime replays the original open input when it restores a
            // panel (for example after an extension reload). Apply each distinct
            // input once so a replay doesn't overwrite later edits.
            const { draftId: _ignored, ...patch } = input;
            if (Object.keys(patch).length) {
                const openInputHash = createHash("sha256").update(JSON.stringify(patch)).digest("base64url");
                if ((await drafts.get(draftId)).openInputHash !== openInputHash) {
                    await updateDraft(draftId, patch, { openInputHash });
                }
            }

            let entry = instances.get(ctx.instanceId);
            if (!entry) {
                const server = await startComposeServer({
                    getState: () => stateFor(entry.draftId),
                    onDraft: async (p) => {
                        await updateDraft(entry.draftId, p);
                        return stateFor(entry.draftId);
                    },
                    onOpen: () => handleOpen(entry.draftId),
                });
                entry = { server, draftId };
                instances.set(ctx.instanceId, entry);
            } else {
                entry.draftId = draftId;
                await entry.server.broadcast();
            }
            return { url: entry.server.url, title: "Teams message" };
        },
        onClose: async (ctx) => {
            const entry = instances.get(ctx.instanceId);
            if (!entry) return;
            instances.delete(ctx.instanceId);
            await entry.server.close();
        },
    });
}
