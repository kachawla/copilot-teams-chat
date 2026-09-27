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

test("choosing a chat clears recipients, and format is normalized", () => {
    const draft = applyDraftPatch(emptyDraft("d"), { to: "a@x.com", format: "html" });
    assert.equal(draft.format, "markdown");
    const withChat = applyDraftPatch(draft, { chatId: "19:x", chatLabel: "X" });
    assert.deepEqual(withChat.to, []);
    assert.equal(applyDraftPatch(withChat, { message: null }).message, "");
});
