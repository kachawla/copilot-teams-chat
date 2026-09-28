import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createService } from "../../com.github.copilot/extensions/teams-chat/lib/service.mjs";

export const URL_A = "https://prod-12.westus.logic.azure.com:443/workflows/abc/triggers/manual/paths/invoke?api-version=2016-06-01&sig=secret";
export const URL_B =
    "https://default1234.ab.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/xyz/triggers/manual/paths/invoke?sig=other";

/** A service with a temp data dir and a fake Teams workflow endpoint that records posts. */
export async function setup({ config = {} } = {}) {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "teams-chat-test-"));
    await fs.writeFile(path.join(dir, "config.json"), JSON.stringify(config));
    // Answers with queued statuses (default 202).
    const webhook = { calls: [], statuses: [] };
    const webhookFetch = async (url, init) => {
        webhook.calls.push({ url, body: JSON.parse(init.body), init });
        const status = webhook.statuses.shift() ?? 202;
        return new Response(null, { status, headers: status === 429 ? { "retry-after": "0" } : {} });
    };
    const service = await createService({ env: { TEAMS_CHAT_DATA_DIR: dir }, webhookFetch });
    const cleanup = () => fs.rm(dir, { recursive: true, force: true });
    return { dir, service, webhook, cleanup };
}

export function fakeSession({ elicitation = true, confirm = true } = {}) {
    const logs = [];
    const prompts = [];
    return {
        logs,
        prompts,
        capabilities: { ui: { elicitation, canvases: true } },
        ui: {
            confirm: async (message) => {
                prompts.push(message);
                return typeof confirm === "function" ? confirm(message) : confirm;
            },
        },
        log: async (message, options) => {
            logs.push({ message, options });
        },
    };
}
