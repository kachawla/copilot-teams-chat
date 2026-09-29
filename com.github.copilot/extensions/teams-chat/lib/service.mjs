import { loadConfig } from "./config.mjs";
import { ContactError, ContactStore } from "./contacts.mjs";
import { WebhookError, WebhookStore } from "./webhooks.mjs";

export async function createService({ env = process.env, webhookFetch = globalThis.fetch } = {}) {
    const config = await loadConfig({ env });
    const webhooks = new WebhookStore(config.dataDir, { extraHostSuffixes: config.allowedWebhookHosts });
    const contacts = new ContactStore(config.dataDir);
    return { config, webhooks, contacts, webhookFetch };
}

export function failure(text) {
    return { textResultForLlm: text, resultType: "failure" };
}

export function describeError(err) {
    if (err instanceof WebhookError || err instanceof ContactError) return err.message;
    return `Unexpected error: ${err?.message || err}`;
}
