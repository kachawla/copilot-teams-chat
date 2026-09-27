import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

// Microsoft Graph Command Line Tools: a Microsoft-published public client that
// supports delegated Graph scopes with user consent. Override with your own
// app registration via config.json or TEAMS_CHAT_CLIENT_ID.
export const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
export const DEFAULT_TENANT = "organizations";
export const DEFAULT_AUTHORITY_HOST = "https://login.microsoftonline.com";
export const DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export const GRAPH_SCOPES = ["ChatMessage.Send", "Chat.ReadBasic", "Chat.Create", "User.ReadBasic.All"];
export const OIDC_SCOPES = ["openid", "profile", "offline_access"];

export function copilotHome(env = process.env) {
    return env.COPILOT_HOME || path.join(os.homedir(), ".copilot");
}

export function dataDir(env = process.env) {
    return env.TEAMS_CHAT_DATA_DIR || path.join(copilotHome(env), "teams-chat");
}

function stripTrailingSlash(value) {
    return String(value).replace(/\/+$/, "");
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

    const config = {
        clientId: env.TEAMS_CHAT_CLIENT_ID || file.clientId || DEFAULT_CLIENT_ID,
        tenantId: env.TEAMS_CHAT_TENANT_ID || file.tenantId || DEFAULT_TENANT,
        authorityHost: stripTrailingSlash(file.authorityHost || DEFAULT_AUTHORITY_HOST),
        graphBaseUrl: stripTrailingSlash(file.graphBaseUrl || DEFAULT_GRAPH_BASE_URL),
        confirmBeforeSend: file.confirmBeforeSend !== false,
        dataDir: dir,
    };

    if (!/^[A-Za-z0-9.-]+$/.test(config.tenantId)) {
        throw new Error(`Invalid tenantId "${config.tenantId}".`);
    }
    if (!/^[0-9a-fA-F-]{36}$/.test(config.clientId)) {
        throw new Error(`Invalid clientId "${config.clientId}". Expected an application (client) ID GUID.`);
    }
    return config;
}
