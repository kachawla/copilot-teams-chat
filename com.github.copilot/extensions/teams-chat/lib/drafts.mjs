import fs from "node:fs/promises";
import path from "node:path";
import { copilotHome } from "./config.mjs";
import { normalizeRecipients } from "./compose.mjs";

const DRAFT_ID = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/;

export function validDraftId(id) {
    return DRAFT_ID.test(String(id));
}

export function emptyDraft(draftId) {
    return {
        draftId,
        to: [],
        topic: null,
        message: "",
        format: "markdown",
        lastResult: null,
        openInputHash: null,
        updatedAt: null,
    };
}

export function applyDraftPatch(draft, patch = {}) {
    const next = { ...draft };
    if (patch.to !== undefined) next.to = normalizeRecipients(patch.to);
    if (patch.topic !== undefined) next.topic = patch.topic ? String(patch.topic) : null;
    if (patch.message !== undefined) next.message = patch.message == null ? "" : String(patch.message);
    if (patch.format !== undefined) next.format = patch.format === "text" ? "text" : "markdown";
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
