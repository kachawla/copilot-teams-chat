import { createHash } from "node:crypto";
import { startComposeServer } from "./compose-server.mjs";
import { deliver, planFingerprint, prepareSend } from "./delivery.mjs";
import { DraftStore, applyDraftPatch, validDraftId } from "./drafts.mjs";
import { describeError } from "./service.mjs";
import { COMPOSE_CANVAS_ID } from "./tools.mjs";

const draftProperties = {
    webhook: { type: "string", description: "Name of the saved Teams workflow to post to." },
    message: { type: "string", description: "Draft message (Markdown unless format is text)." },
    format: { type: "string", enum: ["markdown", "text"] },
};

/**
 * Compose canvas: the agent fills in a draft, the user edits it, previews the
 * formatting, and clicks Send. The agent cannot send from the canvas; only
 * the user's click does.
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
    service.webhooks.onChange(() => broadcastAll());

    const stateFor = async (draftId) => ({
        draft: await drafts.get(draftId),
        webhooks: await service.webhooks.list(),
    });

    const updateDraft = async (draftId, patch, extra = {}) => {
        const draft = { ...applyDraftPatch(await drafts.get(draftId), patch), ...extra };
        await drafts.save(draft);
        await broadcastAll(draftId);
        return drafts.get(draftId);
    };

    const handleSend = async (draftId, { confirm, fingerprint }) => {
        const draft = await drafts.get(draftId);
        try {
            const { plan, payload } = await prepareSend({ service, webhook: draft.webhook, message: draft.message, format: draft.format });
            // The user confirms a specific destination and message. If either
            // changed since (for example the agent updated the draft), ask again.
            const current = planFingerprint(plan, payload);
            if (!confirm || fingerprint !== current) return { needsConfirm: true, label: plan.label, fingerprint: current };
            const result = await deliver({ service, plan, payload });
            await drafts.save({
                ...draft,
                webhook: plan.webhookName,
                message: "",
                lastResult: {
                    ok: true,
                    label: result.label,
                    sentAt: new Date().toISOString(),
                    messagePreview: draft.message.slice(0, 200),
                },
            });
            await broadcastAll(draftId);
            await session.log?.(`Sent a Teams message to ${result.label} from the compose canvas.`);
            return { sent: true, state: await stateFor(draftId) };
        } catch (err) {
            throw new Error(describeError(err));
        }
    };

    const draftIdFor = (instanceId) => instances.get(instanceId)?.draftId ?? "default";

    return createCanvas({
        id: COMPOSE_CANVAS_ID,
        displayName: "Teams message",
        description: "Draft, preview, and send a Microsoft Teams message through a saved workflow; the user reviews it and clicks Send.",
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
                description: "Replace fields of the draft shown in this panel (webhook, message, format).",
                inputSchema: { type: "object", properties: draftProperties },
                handler: async (ctx) => {
                    sessionId ??= ctx.sessionId;
                    return updateDraft(draftIdFor(ctx.instanceId), ctx.input ?? {});
                },
            },
            {
                name: "get_draft",
                description: "Read the current draft, including the user's edits and the result of the last send.",
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
                    onSend: (body) => handleSend(entry.draftId, body),
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
