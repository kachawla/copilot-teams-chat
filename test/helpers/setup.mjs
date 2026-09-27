import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createService } from "../../com.github.copilot/extensions/teams-chat/lib/service.mjs";
import { fakeBrowser, startMockCloud } from "./mock-cloud.mjs";

export async function makeDataDir(config) {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "teams-chat-test-"));
    await fs.writeFile(path.join(dir, "config.json"), JSON.stringify(config));
    return dir;
}

/** Starts a mock cloud and a service pointed at it. `signedIn` completes a browser sign-in first. */
export async function setup({ signedIn = true, config = {}, authOptions = {} } = {}) {
    const cloud = await startMockCloud();
    const dir = await makeDataDir({ authorityHost: cloud.authorityHost, graphBaseUrl: cloud.graphBaseUrl, ...config });
    const opened = [];
    const openUrl = async (url) => {
        opened.push(url);
        return true;
    };
    const env = { TEAMS_CHAT_DATA_DIR: dir };
    const service = await createService({ env, openUrl, authOptions });
    if (signedIn) {
        const pending = await service.auth.startSignIn({ method: "browser" });
        await fakeBrowser(pending.url);
        await service.auth.getAccessToken({ waitForPendingMs: 5000 });
    }
    const cleanup = async () => {
        await service.auth.signOut().catch(() => {});
        await cloud.close();
        await fs.rm(dir, { recursive: true, force: true });
    };
    return { cloud, dir, env, service, opened, openUrl, cleanup };
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
