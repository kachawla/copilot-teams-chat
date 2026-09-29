import assert from "node:assert/strict";
import test from "node:test";
import { MAX_LINK_LENGTH, buildChatLink, markdownToPlainText, toTeamsText } from "../com.github.copilot/extensions/teams-chat/lib/teams-link.mjs";
import { linkParams } from "./helpers/setup.mjs";

test("1:1 link: web and app forms, encoded message, no topic", () => {
    const web = buildChatLink({ emails: ["alice@contoso.com"], topic: "ignored", text: "Hi & bye\nline 2" });
    assert.equal(web, "https://teams.microsoft.com/l/chat/0/0?users=alice%40contoso.com&message=Hi%20%26%20bye%0Aline%202");
    const app = buildChatLink({ emails: ["alice@contoso.com"], text: "Hi", app: true });
    assert.equal(app, "msteams:/l/chat/0/0?users=alice%40contoso.com&message=Hi");
});

test("group link keeps commas between users and includes topic and tenant", () => {
    const link = buildChatLink({ emails: ["a@x.com", "b@x.com"], topic: "Release sync", text: "go", tenantId: "contoso.onmicrosoft.com" });
    assert.match(link, /\?users=a%40x\.com,b%40x\.com&/);
    assert.deepEqual(linkParams(link), { users: "a@x.com,b@x.com", tenantId: "contoso.onmicrosoft.com", topicName: "Release sync", message: "go" });
});

test("links reject empty recipients and overly long messages", () => {
    assert.throws(() => buildChatLink({ emails: [], text: "x" }), /Say who/);
    assert.throws(() => buildChatLink({ emails: ["a@x.com"], text: "x".repeat(MAX_LINK_LENGTH) }), /too long to open in Teams/);
});

test("Markdown becomes readable plain text", () => {
    const md = [
        "# Release **1.2**",
        "",
        "Build is **green** and *fast*, see [the PR](https://x.com/pr/1) or <https://x.com>.",
        "Keep my_var_name and 2*3*4 as is. Use `npm test`.",
        "",
        "- one",
        "  - nested",
        "1. first",
        "> quoted",
        "",
        "---",
        "",
        "```js",
        "const a = **1**;",
        "```",
        "",
        "",
        "",
        "end",
    ].join("\n");
    assert.equal(
        markdownToPlainText(md),
        [
            "Release 1.2",
            "",
            "Build is green and fast, see the PR (https://x.com/pr/1) or https://x.com.",
            "Keep my_var_name and 2*3*4 as is. Use npm test.",
            "",
            "• one",
            "  • nested",
            "1. first",
            "quoted",
            "",
            "const a = **1**;",
            "",
            "end",
        ].join("\n"),
    );
});

test("a link whose label is its URL is not repeated; text format is left alone", () => {
    assert.equal(markdownToPlainText("[https://x.com](https://x.com)"), "https://x.com");
    assert.equal(toTeamsText("  **keep** this\r\n", "text"), "**keep** this");
    assert.throws(() => toTeamsText("   "), /empty/);
    assert.throws(() => toTeamsText("```\n```"), /empty/);
});
