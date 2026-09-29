import assert from "node:assert/strict";
import test from "node:test";
import { openChat, openWithOs } from "../com.github.copilot/extensions/teams-chat/lib/opener.mjs";

test("openChat tries the app first, then the web", async () => {
    const tried = [];
    const openUrl = async (url) => {
        tried.push(url);
        return url === "web";
    };
    assert.equal(await openChat({ appLink: "app", webLink: "web", openUrl }), "web");
    assert.deepEqual(tried, ["app", "web"]);
    assert.equal(await openChat({ appLink: "app", webLink: "web", openIn: "web", openUrl: async () => true }), "web");
    assert.equal(await openChat({ appLink: "app", webLink: "web", openUrl: async () => false }), null);
});

test("openWithOs reports an unhandled URL scheme as not opened", { skip: process.platform === "win32" }, async () => {
    // No app handles this scheme, so open/xdg-open exits non-zero.
    assert.equal(await openWithOs("teams-chat-test-nohandler-7f3a:/x", { timeoutMs: 10_000 }), false);
});
