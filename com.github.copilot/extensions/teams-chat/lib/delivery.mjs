import { createHash } from "node:crypto";
import { isEmail } from "./contacts.mjs";
import { toTeamsHtml } from "./markdown.mjs";
import { MAX_WEBHOOK_PAYLOAD_BYTES, WebhookError, postWebhook, renderWebhookPayload } from "./webhooks.mjs";

// Teams allows 20 people per chat, including you.
export const MAX_RECIPIENTS = 19;

const hashUrl = (url) => createHash("sha256").update(url).digest("base64url");

export function normalizeRecipients(to) {
    const list = Array.isArray(to) ? to : to ? [to] : [];
    return list
        .flatMap((entry) => String(entry).split(/[;\n]/))
        .flatMap((entry) => {
            // Commas separate only lists of addresses, so "Smith, John" stays one name.
            const parts = entry.split(",").map((p) => p.trim()).filter(Boolean);
            return parts.length > 1 && parts.every((p) => p.includes("@")) ? parts : [entry];
        })
        .map((entry) => entry.trim())
        .filter(Boolean);
}

/** Turns emails and saved contact names into people. Never guesses an address. */
export async function resolveRecipients(contacts, to) {
    const entries = normalizeRecipients(to);
    if (!entries.length) throw new WebhookError("Say who to send to: an email address or a saved contact name.", "no_recipients");
    const people = [];
    const unknown = [];
    for (const entry of entries) {
        if (isEmail(entry)) {
            const contact = (await contacts.list()).find((c) => c.email.toLowerCase() === entry.toLowerCase());
            people.push({ name: contact?.name ?? null, email: entry });
            continue;
        }
        const contact = await contacts.find(entry);
        if (contact) people.push({ name: contact.name, email: contact.email });
        else unknown.push(entry);
    }
    if (unknown.length) {
        const saved = await contacts.list();
        throw new WebhookError(
            `No email address is known for ${unknown.map((u) => `"${u}"`).join(", ")}. ` +
                "Ask the user for the address (don't guess it), then send again with the address, or save it first with teams_chat_save_contact." +
                (saved.length ? ` Saved contacts: ${saved.map((c) => c.name).join(", ")}.` : ""),
            "unknown_recipient",
        );
    }
    const unique = people.filter((p, i) => people.findIndex((q) => q.email.toLowerCase() === p.email.toLowerCase()) === i);
    if (unique.length > MAX_RECIPIENTS) {
        throw new WebhookError(`A Teams chat can include at most ${MAX_RECIPIENTS} other people.`, "too_many_recipients");
    }
    return unique;
}

export function formatPerson(person) {
    return person.name ? `${person.name} <${person.email}>` : person.email;
}

async function pickWebhook(webhooks, name, kind) {
    const kindLabel = kind === "direct" ? "send-to-anyone" : "fixed-chat";
    if (name) {
        const entry = await webhooks.get(name);
        if (!entry) {
            throw new WebhookError(`No Teams workflow named "${name}" is saved. Use teams_chat_list_webhooks to see saved ones.`, "not_found");
        }
        if (entry.kind !== kind) {
            throw new WebhookError(
                kind === "direct"
                    ? `"${entry.name}" posts to one fixed chat, so it can't send to specific people. Leave out to, or use a send-to-anyone workflow.`
                    : `"${entry.name}" is a send-to-anyone workflow, so it needs to (who to send to).`,
                "wrong_kind",
            );
        }
        return entry;
    }
    const saved = (await webhooks.list()).filter((w) => w.kind === kind);
    if (saved.length === 1) return webhooks.get(saved[0].name);
    if (!saved.length) {
        throw new WebhookError(
            kind === "direct"
                ? 'No send-to-anyone workflow is saved, so messages can\'t go to specific people yet. Point the user to "Send to anyone" in the README to create one, then add it with teams_chat_add_webhook and kind: "direct".'
                : "No Teams workflow webhooks are saved yet. Ask the user to create one and add it with teams_chat_add_webhook.",
            "none_saved",
        );
    }
    throw new WebhookError(`Several ${kindLabel} Teams workflows are saved (${saved.map((w) => w.name).join(", ")}). Say which one with webhook.`, "ambiguous");
}

function renderDirectPayload({ people, topic, message, format }) {
    const text = String(message ?? "");
    if (!text.trim()) throw new WebhookError("The message is empty.", "empty");
    const payload = {
        to: people.map((p) => p.email).join(";"),
        topic: people.length > 1 ? String(topic || "").trim().slice(0, 250) || null : null,
        html: toTeamsHtml(text, format === "text" ? "text" : "markdown"),
        text,
    };
    const bytes = new TextEncoder().encode(JSON.stringify(payload)).length;
    if (bytes > MAX_WEBHOOK_PAYLOAD_BYTES) {
        throw new WebhookError(`The message is too long for Teams (${bytes} bytes, limit ${MAX_WEBHOOK_PAYLOAD_BYTES}). Shorten it or split it.`, "too_long");
    }
    return payload;
}

/**
 * Resolves the destination and renders the message without sending anything.
 * With `to`, uses a send-to-anyone ("direct") workflow; otherwise a fixed-chat one.
 */
export async function prepareSend({ service, webhook, to, topic, message, format }) {
    const wantsPeople = normalizeRecipients(to).length > 0;
    const entry = await pickWebhook(service.webhooks, webhook, wantsPeople ? "direct" : "chat");
    if (!wantsPeople) {
        return {
            plan: { kind: "chat", webhookName: entry.name, urlHash: hashUrl(entry.url), label: `the "${entry.name}" Teams workflow`, people: [] },
            payload: renderWebhookPayload(message, format),
        };
    }
    const people = await resolveRecipients(service.contacts, to);
    const payload = renderDirectPayload({ people, topic, message, format });
    const who = people.map(formatPerson).join(", ");
    const label =
        people.length === 1
            ? `${who} (1:1 chat, as you)`
            : `a new group chat with ${who}${payload.topic ? ` named "${payload.topic}"` : ""} (as you)`;
    return {
        plan: { kind: "direct", webhookName: entry.name, urlHash: hashUrl(entry.url), label, people, topic: payload.topic },
        payload,
    };
}

export async function deliver({ service, plan, payload }) {
    const entry = await service.webhooks.get(plan.webhookName);
    if (!entry || hashUrl(entry.url) !== plan.urlHash || entry.kind !== plan.kind) {
        throw new WebhookError(`The "${plan.webhookName}" workflow changed or was removed. Nothing was sent.`, "changed");
    }
    await postWebhook(entry.url, payload, { fetchImpl: service.webhookFetch });
    return { label: plan.label, kind: plan.kind };
}

/** Identifies exactly what a confirmed send will do: destination and rendered content. */
export function planFingerprint(plan, payload) {
    return createHash("sha256")
        .update(JSON.stringify([plan.kind, plan.webhookName, plan.urlHash, plan.people.map((p) => p.email.toLowerCase())]))
        .update("\0")
        .update(JSON.stringify(payload))
        .digest("base64url");
}

export function confirmationText(plan, message, maxChars = 1200) {
    const body = String(message);
    const preview = body.length > maxChars ? `${body.slice(0, maxChars)}\n… (${body.length - maxChars} more characters)` : body;
    return `Send this Microsoft Teams message to ${plan.label}?\n\n${preview}`;
}
