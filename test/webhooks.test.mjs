import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { createTools } from "../com.github.copilot/extensions/teams-chat/lib/tools.mjs";
import { buildAdaptiveCard, validateWebhookUrl } from "../com.github.copilot/extensions/teams-chat/lib/webhooks.mjs";
import { fakeSession, setup } from "./helpers/setup.mjs";

const URL_A = "https://prod-12.westus.logic.azure.com:443/workflows/abc/triggers/manual/paths/invoke?api-version=2016-06-01&sig=secret";
const URL_B = "https://default1234.ab.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/xyz/triggers/manual/paths/invoke?sig=other";

function toolsFor(service, session) {
    return Object.fromEntries(createTools({ service, session }).map((tool) => [tool.name, tool]));
}

test("only Teams Workflows webhook hosts are accepted", () => {
    assert.ok(validateWebhookUrl(URL_A));
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
    assert.equal(contentType, "application/vnd.microsoft.card.adaptive");
    assert.equal(content.type, "AdaptiveCard");
    assert.deepEqual(content.body, [
        { type: "TextBlock", text: "Release", wrap: true, weight: "Bolder", size: "Medium" },
        { type: "TextBlock", text: "**green** and [PR](https://x.com)", wrap: true },
        { type: "TextBlock", text: "npm test", wrap: true, fontType: "Monospace" },
        { type: "TextBlock", text: "- a\n- b", wrap: true },
    ]);
    assert.throws(() => buildAdaptiveCard("   "), /empty/);
});

test("adding a webhook asks the user, stores it privately, and never shows the URL", async (t) => {
    const { service, dir, cleanup } = await setup({ signedIn: false });
    t.after(cleanup);
    const session = fakeSession({ confirm: true });
    const tools = toolsFor(service, session);
    const added = await tools.teams_chat_add_webhook.handler({ name: "Team chat", url: URL_A, description: "our team" });
    assert.match(added, /Saved the Teams workflow webhook "Team chat"/);
    assert.match(session.prompts[0], /Save the Teams workflow webhook "Team chat" \(prod-12\.westus\.logic\.azure\.com\)/);
    const listed = await tools.teams_chat_list_webhooks.handler({});
    assert.equal(listed, "- Team chat: our team");
    assert.doesNotMatch(listed + (await tools.teams_chat_status.handler({})), /sig=secret/);
    const file = path.join(dir, "webhooks.json");
    if (process.platform !== "win32") assert.equal((await fs.stat(file)).mode & 0o777, 0o600);
    assert.equal(JSON.parse(await fs.readFile(file, "utf8")).webhooks["Team chat"].url, URL_A);

    assert.match(await tools.teams_chat_remove_webhook.handler({ name: "team chat" }), /Removed/);
    assert.match(await tools.teams_chat_list_webhooks.handler({}), /No Teams workflow webhooks/);
});

test("declined or untrusted webhooks are not saved", async (t) => {
    const { service, cleanup } = await setup({ signedIn: false });
    t.after(cleanup);
    const declined = await toolsFor(service, fakeSession({ confirm: false })).teams_chat_add_webhook.handler({ name: "x", url: URL_A });
    assert.equal(declined.resultType, "rejected");
    const untrusted = await toolsFor(service, fakeSession()).teams_chat_add_webhook.handler({ name: "x", url: "https://attacker.example.com/hook" });
    assert.equal(untrusted.resultType, "failure");
    assert.match(untrusted.textResultForLlm, /not a Teams Workflows webhook host/);
    const noUi = await toolsFor(service, fakeSession({ elicitation: false })).teams_chat_add_webhook.handler({ name: "x", url: URL_A });
    assert.equal(noUi.resultType, "failure");
    assert.deepEqual(await service.webhooks.list(), []);
});

test("sending to a webhook needs no sign-in, confirms, and posts an adaptive card", async (t) => {
    const { service, webhook, cloud, cleanup } = await setup({ signedIn: false });
    t.after(cleanup);
    await service.webhooks.add("Team chat", URL_A);
    const session = fakeSession();
    const result = await toolsFor(service, session).teams_chat_send_message.handler({ webhook: "team chat", message: "Build is **green**" });
    assert.equal(result, 'Sent to the "Team chat" Teams workflow (posts as the Workflows bot).');
    assert.match(session.prompts[0], /Send this Microsoft Teams message to the "Team chat" Teams workflow/);
    assert.equal(webhook.calls.length, 1);
    assert.equal(webhook.calls[0].url, URL_A);
    assert.equal(webhook.calls[0].body.attachments[0].content.body[0].text, "Build is **green**");
    assert.equal(cloud.state.requests.length, 0, "no Graph or sign-in traffic");
});

test("webhook sends: declined, unknown name, throttled, and server errors", async (t) => {
    const { service, webhook, cleanup } = await setup({ signedIn: false });
    t.after(cleanup);
    await service.webhooks.add("chat", URL_A);

    const declined = await toolsFor(service, fakeSession({ confirm: false })).teams_chat_send_message.handler({ webhook: "chat", message: "x" });
    assert.equal(declined.resultType, "rejected");
    assert.equal(webhook.calls.length, 0);

    const tools = toolsFor(service, fakeSession());
    const unknown = await tools.teams_chat_send_message.handler({ webhook: "nope", message: "x" });
    assert.match(unknown.textResultForLlm, /No Teams workflow named "nope"/);

    webhook.statuses.push(429);
    assert.match(await tools.teams_chat_send_message.handler({ webhook: "chat", message: "x" }), /^Sent/);
    assert.equal(webhook.calls.length, 2, "429 is retried once");

    webhook.statuses.push(502);
    const failed = await tools.teams_chat_send_message.handler({ webhook: "chat", message: "x" });
    assert.match(failed.textResultForLlm, /may still have been posted; check Teams/);
    assert.equal(webhook.calls.length, 3, "5xx is not retried");

    webhook.statuses.push(404);
    const gone = await tools.teams_chat_send_message.handler({ webhook: "chat", message: "x" });
    assert.match(gone.textResultForLlm, /workflow may be turned off, deleted/);
});

test("a webhook whose URL changes after confirmation is not used", async (t) => {
    const { service, webhook, cleanup } = await setup({ signedIn: false });
    t.after(cleanup);
    await service.webhooks.add("chat", URL_A);
    const session = fakeSession({
        confirm: async () => {
            await service.webhooks.add("chat", URL_B);
            return true;
        },
    });
    const result = await toolsFor(service, session).teams_chat_send_message.handler({ webhook: "chat", message: "x" });
    assert.match(result.textResultForLlm, /changed or was removed. Nothing was sent/);
    assert.equal(webhook.calls.length, 0);
});
