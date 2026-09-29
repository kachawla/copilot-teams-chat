import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

export function copilotHome(env = process.env) {
    return env.COPILOT_HOME || path.join(os.homedir(), ".copilot");
}

export function dataDir(env = process.env) {
    return env.TEAMS_CHAT_DATA_DIR || path.join(copilotHome(env), "teams-chat");
}

export async function loadConfig({ env = process.env, dir = dataDir(env) } = {}) {
    let file = {};
    try {
        file = JSON.parse(await fs.readFile(path.join(dir, "config.json"), "utf8"));
    } catch (err) {
        if (err.code !== "ENOENT") {
            throw new Error(`Could not read ${path.join(dir, "config.json")}: ${err.message}`);
        }
    }
    const tenantId = file.tenantId ? String(file.tenantId) : null;
    if (tenantId && !/^[A-Za-z0-9.-]+$/.test(tenantId)) throw new Error(`Invalid tenantId "${tenantId}".`);
    return {
        openIn: file.openIn === "web" ? "web" : "app",
        tenantId,
        dataDir: dir,
    };
}
