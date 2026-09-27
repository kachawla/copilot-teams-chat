import { createHash } from "node:crypto";
import { RecipientError, executeSend, planSend, renderMessage } from "./send.mjs";
import { requireAccount } from "./service.mjs";
import { postWebhook, renderWebhookPayload } from "./webhooks.mjs";

/**
 * Resolves where a message will go and renders it, without sending anything.
 * `target` has either `webhook` (a saved workflow name) or `chatId`/`to`/`topic`.
 */
export async function prepareSend({ service, target = {}, message, format, waitForSignInMs = 0 }) {
    if (target.webhook) {
        const entry = await service.webhooks.get(target.webhook);
        if (!entry) {
            throw new RecipientError(`No Teams workflow named "${target.webhook}" is set up. Use teams_chat_list_webhooks to see saved ones.`);
        }
        const plan = {
            kind: "webhook",
            webhookName: entry.name,
            urlHash: createHash("sha256").update(entry.url).digest("base64url"),
            label: `the "${entry.name}" Teams workflow (posts as the Workflows bot)`,
            people: [],
        };
        return { plan, payload: renderWebhookPayload(message, format), account: null };
    }
    const account = await requireAccount(service.auth, waitForSignInMs);
    const html = renderMessage(message, format);
    const plan = await planSend({ graph: service.graph, account, chatId: target.chatId, to: target.to, topic: target.topic });
    return { plan, payload: html, account };
}

export async function deliver({ service, plan, payload, account }) {
    if (plan.kind === "webhook") {
        const entry = await service.webhooks.get(plan.webhookName);
        const urlHash = entry && createHash("sha256").update(entry.url).digest("base64url");
        if (!entry || urlHash !== plan.urlHash) {
            throw new RecipientError(`The "${plan.webhookName}" workflow changed or was removed. Nothing was sent.`);
        }
        await postWebhook(entry.url, payload, { fetchImpl: service.webhookFetch });
        return { chatId: null, messageId: null, webUrl: null, label: plan.label };
    }
    return executeSend({ graph: service.graph, account, plan, html: payload });
}

/** Identifies exactly what a confirmed send will do: target and rendered content. */
export function planFingerprint(plan, payload) {
    const target = {
        kind: plan.kind,
        chatId: plan.chatId ?? null,
        people: (plan.people ?? []).map((p) => p.id).sort(),
        topic: plan.topic ?? null,
        webhook: plan.kind === "webhook" ? [plan.webhookName, plan.urlHash] : null,
    };
    const content = typeof payload === "string" ? payload : JSON.stringify(payload);
    return createHash("sha256").update(JSON.stringify(target)).update("\0").update(content).digest("base64url");
}
