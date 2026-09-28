import assert from "node:assert/strict";
import path from "node:path";
import test from "node:test";
import { applyDraftPatch, createDraftsDirResolver, emptyDraft } from "../com.github.copilot/extensions/teams-chat/lib/drafts.mjs";

test("drafts dir is derived from the session ID before the session is joined, then stays fixed", () => {
    let workspacePath;
    const resolve = createDraftsDirResolver({ getWorkspacePath: () => workspacePath, fallbackDir: "/data", env: { COPILOT_HOME: "/home/.copilot" } });
    assert.equal(resolve(undefined), path.join("/data", "drafts"), "no session yet: not cached");
    const early = resolve("abc");
    assert.equal(early, path.join("/home/.copilot", "session-state", "abc", "files", "teams-chat-drafts"));
    workspacePath = "/somewhere/else";
    assert.equal(resolve("abc"), early, "does not switch once chosen");
});

test("drafts dir prefers the session workspace path when known", () => {
    const resolve = createDraftsDirResolver({ getWorkspacePath: () => "/ws", fallbackDir: "/data" });
    assert.equal(resolve("abc"), path.join("/ws", "files", "teams-chat-drafts"));
});

test("draft patches normalize values", () => {
    const draft = applyDraftPatch(emptyDraft("d"), { webhook: "team chat", message: null, format: "html", ignored: "x" });
    assert.equal(draft.webhook, "team chat");
    assert.equal(draft.message, "");
    assert.equal(draft.format, "markdown");
    assert.equal(draft.ignored, undefined);
    assert.equal(applyDraftPatch(draft, { webhook: "" }).webhook, null);
});
