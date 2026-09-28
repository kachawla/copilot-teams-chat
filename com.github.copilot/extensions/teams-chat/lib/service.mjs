import { loadConfig } from "./config.mjs";
import { WebhookError, WebhookStore } from "./webhooks.mjs";

export async function createService({ env = process.env, webhookFetch = globalThis.fetch } = {}) {
    const config = await loadConfig({ env });
    const webhooks = new WebhookStore(config.dataDir, { extraHostSuffixes: config.allowedWebhookHosts });
    return { config, webhooks, webhookFetch };
}

export function failure(text) {
    return { textResultForLlm: text, resultType: "failure" };
}

export function describeError(err) {
    if (err instanceof WebhookError) return err.message;
    return `Unexpected error: ${err?.message || err}`;
}
