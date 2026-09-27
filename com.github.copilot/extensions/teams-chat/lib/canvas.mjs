import { createHash } from "node:crypto";
import { startComposeServer } from "./compose-server.mjs";
import { DraftStore, applyDraftPatch, validDraftId } from "./drafts.mjs";
import { executeSend, planFingerprint, planSend, renderMessage } from "./send.mjs";
import { describeError, requireAccount } from "./service.mjs";
import { COMPOSE_CANVAS_ID } from "./tools.mjs";

const draftProperties = {
    to: { type: "array", items: { type: "string" }, description: 'Recipients: email addresses or full names; "me" for Notes to self.' },
    chatId: { type: "string", description: "Existing chat ID from teams_chat_list_chats. Replaces to." },
    chatLabel: { type: "string", description: "Display name for chatId, shown to the user." },
    topic: { type: "string", description: "Name for a new group chat, if one has to be created." },
    message: { type: "string", description: "Draft message (Markdown unless format is text)." },
    format: { type: "string", enum: ["markdown", "text"] },
};

/**
 * Compose canvas: the agent fills in a draft, the user edits it, previews the
 * Teams formatting, and clicks Send. The agent cannot send from the canvas;
 * only the user's click does.
 */
export function createComposeCanvas({ createCanvas, CanvasError, service, session, draftsDir }) {
    const { auth, graph } = service;
    let sessionId;
    const drafts = new DraftStore(() => draftsDir(sessionId));
    const instances = new Map(); // instanceId -> { server, draftId }

    const broadcastAll = async (draftId) => {
        for (const entry of instances.values()) {
            if (!draftId || entry.draftId === draftId) await entry.server.broadcast().catch(() => {});
        }
    };
    auth.on("change", () => broadcastAll());

    const stateFor = async (draftId) => ({ auth: await auth.status(), draft: await drafts.get(draftId) });

    const updateDraft = async (draftId, patch, extra = {}) => {
        const draft = { ...applyDraftPatch(await drafts.get(draftId), patch), ...extra };
        await drafts.save(draft);
        await broadcastAll(draftId);
        return drafts.get(draftId);
    };

    const handleSend = async (draftId, { confirm, fingerprint }) => {
        let account;
        try {
            account = await requireAccount(auth);
        } catch (err) {
            throw new Error(describeError(err));
        }
        const draft = await drafts.get(draftId);
        try {
            const html = renderMessage(draft.message, draft.format);
            const plan = await planSend({ graph, account, chatId: draft.chatId, to: draft.to, topic: draft.topic });
            // The user confirms a specific target and message. If either changed
            // since (for example the agent updated the draft), ask again.
            const current = planFingerprint(plan, html);
            if (!confirm || fingerprint !== current) return { needsConfirm: true, label: plan.label, fingerprint: current };
            const result = await executeSend({ graph, account, plan, html });
            const sent = {
                ...draft,
                message: "",
                lastResult: {
                    ok: true,
                    label: plan.label,
                    chatId: result.chatId,
                    messageId: result.messageId,
                    webUrl: result.webUrl,
                    sentAt: new Date().toISOString(),
                    messagePreview: draft.message.slice(0, 200),
                },
            };
            if (plan.kind === "oneOnOne" || plan.kind === "newGroup") {
                sent.chatId = result.chatId;
                sent.chatLabel = plan.label.replace(/^(a new |1:1 chat with )/, "");
                sent.to = [];
            }
            await drafts.save(sent);
            await broadcastAll(draftId);
            await session.log?.(`Sent a Teams message to ${plan.label} from the compose canvas.`);
            return { sent: true, state: await stateFor(draftId) };
        } catch (err) {
            throw new Error(describeError(err));
        }
    };

    const draftIdFor = (instanceId) => instances.get(instanceId)?.draftId ?? "default";

    return createCanvas({
        id: COMPOSE_CANVAS_ID,
        displayName: "Teams message",
        description: "Draft, preview, and send a Microsoft Teams chat message; the user reviews it and clicks Send.",
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
                description: "Replace fields of the draft shown in this panel (recipients, chat, topic, message, format).",
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
                    onSignIn: async (method) => {
                        await auth.startSignIn({ method, force: true });
                        return stateFor(entry.draftId);
                    },
                    onSignOut: async () => {
                        await auth.signOut();
                        return stateFor(entry.draftId);
                    },
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
