import { createHash } from "node:crypto";
import { WebhookError, postWebhook, renderWebhookPayload } from "./webhooks.mjs";

const hashUrl = (url) => createHash("sha256").update(url).digest("base64url");

/** Picks the saved webhook. With no name given, uses the only one if exactly one is saved. */
export async function resolveWebhook(webhooks, name) {
    if (name) {
        const entry = await webhooks.get(name);
        if (!entry) {
            throw new WebhookError(`No Teams workflow named "${name}" is saved. Use teams_chat_list_webhooks to see saved ones.`, "not_found");
        }
        return entry;
    }
    const saved = await webhooks.list();
    if (saved.length === 1) return webhooks.get(saved[0].name);
    if (!saved.length) {
        throw new WebhookError("No Teams workflow webhooks are saved yet. Ask the user to create one and add it with teams_chat_add_webhook.", "none_saved");
    }
    throw new WebhookError(`Several Teams workflows are saved (${saved.map((w) => w.name).join(", ")}). Say which one with webhook.`, "ambiguous");
}

/** Resolves the destination and renders the message without sending anything. */
export async function prepareSend({ service, webhook, message, format }) {
    const entry = await resolveWebhook(service.webhooks, webhook);
    const plan = {
        webhookName: entry.name,
        urlHash: hashUrl(entry.url),
        label: `the "${entry.name}" Teams workflow`,
    };
    return { plan, payload: renderWebhookPayload(message, format) };
}

export async function deliver({ service, plan, payload }) {
    const entry = await service.webhooks.get(plan.webhookName);
    if (!entry || hashUrl(entry.url) !== plan.urlHash) {
        throw new WebhookError(`The "${plan.webhookName}" workflow changed or was removed. Nothing was sent.`, "changed");
    }
    await postWebhook(entry.url, payload, { fetchImpl: service.webhookFetch });
    return { label: plan.label };
}

/** Identifies exactly what a confirmed send will do: destination and rendered content. */
export function planFingerprint(plan, payload) {
    return createHash("sha256")
        .update(JSON.stringify([plan.webhookName, plan.urlHash]))
        .update("\0")
        .update(JSON.stringify(payload))
        .digest("base64url");
}

export function confirmationText(plan, message, maxChars = 1200) {
    const body = String(message);
    const preview = body.length > maxChars ? `${body.slice(0, maxChars)}\n… (${body.length - maxChars} more characters)` : body;
    return `Send this message to ${plan.label} in Microsoft Teams?\n\n${preview}`;
}
