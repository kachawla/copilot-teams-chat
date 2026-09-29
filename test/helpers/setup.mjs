import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createService } from "../../com.github.copilot/extensions/teams-chat/lib/service.mjs";

/**
 * A service with a temp data dir and a fake OS opener that records URLs.
 * `handles` decides which URLs "open" (default: all).
 */
export async function setup({ config = {}, handles = () => true } = {}) {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "teams-chat-test-"));
    await fs.writeFile(path.join(dir, "config.json"), JSON.stringify(config));
    const opened = [];
    const openUrl = async (url) => {
        opened.push(url);
        return handles(url);
    };
    const service = await createService({ env: { TEAMS_CHAT_DATA_DIR: dir }, openUrl });
    const cleanup = () => fs.rm(dir, { recursive: true, force: true });
    return { dir, service, opened, cleanup };
}

export function fakeSession() {
    const logs = [];
    return {
        logs,
        log: async (message, options) => {
            logs.push({ message, options });
        },
    };
}

export function linkParams(link) {
    const query = link.slice(link.indexOf("?") + 1);
    return Object.fromEntries(
        query.split("&").map((pair) => {
            const [k, v] = pair.split("=");
            return [k, decodeURIComponent(v)];
        }),
    );
}
