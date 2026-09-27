import fs from "node:fs/promises";
import path from "node:path";
import { copilotHome } from "./config.mjs";
import { splitRecipients } from "./recipients.mjs";

const DRAFT_ID = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/;

export function validDraftId(id) {
    return DRAFT_ID.test(String(id));
}

export function emptyDraft(draftId) {
    return {
        draftId,
        to: [],
        chatId: null,
        chatLabel: null,
        webhook: null,
        topic: null,
        message: "",
        format: "markdown",
        lastResult: null,
        openInputHash: null,
        updatedAt: null,
    };
}

const EDITABLE = ["to", "chatId", "chatLabel", "webhook", "topic", "message", "format"];

export function applyDraftPatch(draft, patch = {}) {
    const next = { ...draft };
    for (const key of EDITABLE) {
        if (patch[key] === undefined) continue;
        if (key === "to") {
            next.to = splitRecipients(patch.to);
        } else if (key === "format") {
            next.format = patch.format === "text" ? "text" : "markdown";
        } else {
            next[key] = patch[key] === null || patch[key] === "" ? (key === "message" ? "" : null) : String(patch[key]);
        }
    }
    // Choosing recipients, a chat, or a webhook replaces the other targets.
    if (patch.to !== undefined && patch.chatId === undefined && next.to.length) {
        next.chatId = null;
        next.chatLabel = null;
    }
    if (patch.chatId && patch.to === undefined) next.to = [];
    if (patch.webhook) {
        next.to = [];
        next.chatId = null;
        next.chatLabel = null;
    } else if ((patch.to !== undefined && next.to.length) || patch.chatId) {
        next.webhook = null;
    }
    return next;
}

/**
 * Drafts live in the session workspace. The runtime can call canvas open()
 * before joinSession resolves (to restore panels), when workspacePath is not
 * known yet, so fall back to the standard session-state path for the request's
 * session ID. The first answer is kept so reads and writes never switch
 * directories mid-session.
 */
export function createDraftsDirResolver({ getWorkspacePath, fallbackDir, env = process.env }) {
    let resolved;
    return (sessionId) => {
        if (resolved) return resolved;
        const workspace = getWorkspacePath() || (sessionId ? path.join(copilotHome(env), "session-state", sessionId) : null);
        if (!workspace) return path.join(fallbackDir, "drafts");
        resolved = path.join(workspace, "files", "teams-chat-drafts");
        return resolved;
    };
}

// Drafts are per Copilot session, stored as files in the session workspace so
// they survive canvas reloads and extension restarts.
export class DraftStore {
    /** @param {string | (() => string)} dir directory, or a function resolved on each access */
    constructor(dir) {
        this.dir = dir;
    }

    get path() {
        return typeof this.dir === "function" ? this.dir() : this.dir;
    }

    #file(draftId) {
        if (!validDraftId(draftId)) throw new Error(`Invalid draftId "${draftId}".`);
        return path.join(this.path, `${draftId}.json`);
    }

    async get(draftId) {
        try {
            return { ...emptyDraft(draftId), ...JSON.parse(await fs.readFile(this.#file(draftId), "utf8")), draftId };
        } catch (err) {
            if (err.code === "ENOENT" || err instanceof SyntaxError) return emptyDraft(draftId);
            throw err;
        }
    }

    async save(draft) {
        await fs.mkdir(this.path, { recursive: true, mode: 0o700 });
        const saved = { ...draft, updatedAt: new Date().toISOString() };
        await fs.writeFile(this.#file(draft.draftId), JSON.stringify(saved, null, 2), { mode: 0o600 });
        return saved;
    }
}
