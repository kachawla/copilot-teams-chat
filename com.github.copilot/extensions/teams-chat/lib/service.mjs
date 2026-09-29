import { loadConfig } from "./config.mjs";
import { ContactError, ContactStore } from "./contacts.mjs";
import { openWithOs } from "./opener.mjs";
import { LinkError } from "./teams-link.mjs";

export async function createService({ env = process.env, openUrl = openWithOs } = {}) {
    const config = await loadConfig({ env });
    return { config, contacts: new ContactStore(config.dataDir), openUrl };
}

export function failure(text) {
    return { textResultForLlm: text, resultType: "failure" };
}

export function describeError(err) {
    if (err instanceof ContactError || err instanceof LinkError) return err.message;
    return `Unexpected error: ${err?.message || err}`;
}
