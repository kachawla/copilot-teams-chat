import assert from "node:assert/strict";
import test from "node:test";
import { buildAdaptiveCard, postWebhook, validateWebhookUrl } from "../com.github.copilot/extensions/teams-chat/lib/webhooks.mjs";
import { URL_A, URL_B } from "./helpers/setup.mjs";

test("only Teams Workflows webhook hosts are accepted", () => {
    assert.equal(validateWebhookUrl(`  ${URL_A} `), URL_A, "kept exactly, trimmed");
    assert.ok(validateWebhookUrl(URL_B));
    assert.ok(validateWebhookUrl("https://contoso.webhook.office.com/webhookb2/x"));
    for (const bad of [
        "http://prod-12.westus.logic.azure.com/workflows/x",
        "https://evil.example.com/logic.azure.com",
        "https://logic.azure.com.evil.example.com/x",
        "https://evillogic.azure.com/x",
        "https://user:pw@prod-1.westus.logic.azure.com/x",
        "not a url",
    ]) {
        assert.throws(() => validateWebhookUrl(bad), undefined, bad);
    }
    assert.ok(validateWebhookUrl("https://hooks.contoso.example/x", ["contoso.example"]), "extra hosts come from config");
});

test("adaptive card keeps Markdown text and styles code and headings", () => {
    const card = buildAdaptiveCard("# Release\n**green** and [PR](https://x.com)\n\n```\nnpm test\n```\n- a\n- b");
    const { content, contentType } = card.attachments[0];
    assert.equal(card.type, "message");
    assert.equal(contentType, "application/vnd.microsoft.card.adaptive");
    assert.equal(content.type, "AdaptiveCard");
    assert.deepEqual(content.body, [
        { type: "TextBlock", text: "Release", wrap: true, weight: "Bolder", size: "Medium" },
        { type: "TextBlock", text: "**green** and [PR](https://x.com)", wrap: true },
        { type: "TextBlock", text: "npm test", wrap: true, fontType: "Monospace" },
        { type: "TextBlock", text: "- a\n- b", wrap: true },
    ]);
    assert.deepEqual(buildAdaptiveCard("# not a heading", "text").attachments[0].content.body, [
        { type: "TextBlock", text: "# not a heading", wrap: true },
    ]);
    assert.throws(() => buildAdaptiveCard("   "), /empty/);
});

test("postWebhook reports network errors without leaking the URL", async () => {
    const fetchImpl = async () => {
        throw new Error("getaddrinfo ENOTFOUND");
    };
    await assert.rejects(postWebhook(URL_A, {}, { fetchImpl }), (err) => /Could not reach the Teams webhook/.test(err.message) && !err.message.includes("sig="));
});
