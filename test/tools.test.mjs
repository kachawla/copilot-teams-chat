import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { createTools } from "../com.github.copilot/extensions/teams-chat/lib/tools.mjs";
import { URL_A, URL_B, fakeSession, setup } from "./helpers/setup.mjs";

function toolsFor(service, session) {
    return Object.fromEntries(createTools({ service, session }).map((tool) => [tool.name, tool]));
}

test("tool names are namespaced and only read-only tools and send skip the permission prompt", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const tools = createTools({ service, session: fakeSession() });
    assert.deepEqual(tools.map((tool) => tool.name), [
        "teams_chat_status",
        "teams_chat_send_message",
        "teams_chat_add_webhook",
        "teams_chat_list_webhooks",
        "teams_chat_remove_webhook",
        "teams_chat_save_contact",
        "teams_chat_list_contacts",
        "teams_chat_remove_contact",
    ]);
    const skip = tools.filter((tool) => tool.skipPermission).map((tool) => tool.name).sort();
    assert.deepEqual(skip, ["teams_chat_list_contacts", "teams_chat_list_webhooks", "teams_chat_send_message", "teams_chat_status"]);
});

test("status explains setup when nothing is saved", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession());
    assert.match(await tools.teams_chat_status.handler({}), /No Teams workflows saved yet\. For one fixed chat or channel: in Teams, open Apps > Workflows.*"Send to anyone"/);
    await service.webhooks.add("team chat", URL_A);
    assert.match(await tools.teams_chat_status.handler({}), /Saved Teams workflows: team chat \(one fixed chat, as the Workflows bot\)\./);
    assert.match(await toolsFor(service, fakeSession({ elicitation: false })).teams_chat_status.handler({}), /can't show confirmation prompts/);
});

test("adding a webhook asks the user, stores it privately, and never shows the URL", async (t) => {
    const { service, dir, cleanup } = await setup();
    t.after(cleanup);
    const session = fakeSession();
    const tools = toolsFor(service, session);
    const added = await tools.teams_chat_add_webhook.handler({ name: "Team chat", url: URL_A, description: "our team" });
    assert.match(added, /Saved the Teams workflow webhook "Team chat"/);
    assert.match(session.prompts[0], /Save the Teams workflow webhook "Team chat" \(prod-12\.westus\.logic\.azure\.com\)/);
    const listed = await tools.teams_chat_list_webhooks.handler({});
    assert.equal(listed, "- Team chat (fixed chat): our team");
    assert.doesNotMatch(listed + (await tools.teams_chat_status.handler({})), /sig=secret/);
    const file = path.join(dir, "webhooks.json");
    if (process.platform !== "win32") assert.equal((await fs.stat(file)).mode & 0o777, 0o600);
    assert.equal(JSON.parse(await fs.readFile(file, "utf8")).webhooks["Team chat"].url, URL_A, "URL is stored exactly as given");

    assert.match(await tools.teams_chat_add_webhook.handler({ name: "team CHAT", url: URL_B }), /Replaced/);
    assert.equal((await service.webhooks.list()).length, 1);
    assert.match(await tools.teams_chat_remove_webhook.handler({ name: "team chat" }), /Removed/);
    assert.match(await tools.teams_chat_list_webhooks.handler({}), /No Teams workflow webhooks/);
    assert.equal((await tools.teams_chat_remove_webhook.handler({ name: "gone" })).resultType, "failure");
});

test("declined, invalid, or untrusted webhooks are not saved", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const declined = await toolsFor(service, fakeSession({ confirm: false })).teams_chat_add_webhook.handler({ name: "x", url: URL_A });
    assert.equal(declined.resultType, "rejected");
    const tools = toolsFor(service, fakeSession());
    const untrusted = await tools.teams_chat_add_webhook.handler({ name: "x", url: "https://attacker.example.com/hook" });
    assert.match(untrusted.textResultForLlm, /not a Teams Workflows webhook host/);
    assert.match((await tools.teams_chat_add_webhook.handler({ name: "x", url: "nope" })).textResultForLlm, /not a valid URL/);
    assert.match((await tools.teams_chat_add_webhook.handler({ name: "../x", url: URL_A })).textResultForLlm, /Invalid webhook name/);
    const noUi = await toolsFor(service, fakeSession({ elicitation: false })).teams_chat_add_webhook.handler({ name: "x", url: URL_A });
    assert.equal(noUi.resultType, "failure");
    assert.deepEqual(await service.webhooks.list(), []);
});

test("send confirms, then posts an adaptive card", async (t) => {
    const { service, webhook, cleanup } = await setup();
    t.after(cleanup);
    await service.webhooks.add("Team chat", URL_A);
    const session = fakeSession();
    const result = await toolsFor(service, session).teams_chat_send_message.handler({ webhook: "team chat", message: "Build is **green**" });
    assert.equal(result, 'Sent to the "Team chat" Teams workflow. It appears from the Workflows bot.');
    assert.match(session.prompts[0], /^Send this Microsoft Teams message to the "Team chat" Teams workflow\?\n\nBuild is \*\*green\*\*$/);
    assert.equal(webhook.calls.length, 1);
    assert.equal(webhook.calls[0].url, URL_A);
    assert.equal(webhook.calls[0].init.redirect, "error");
    assert.equal(webhook.calls[0].body.attachments[0].content.body[0].text, "Build is **green**");
    assert.equal(session.logs[0].message, 'Sent a Teams message to the "Team chat" Teams workflow.');
});

test("send picks the only saved workflow, and asks which one when there are several", async (t) => {
    const { service, webhook, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession());
    assert.match((await tools.teams_chat_send_message.handler({ message: "x" })).textResultForLlm, /No Teams workflow webhooks are saved yet/);
    await service.webhooks.add("one", URL_A);
    assert.match(await tools.teams_chat_send_message.handler({ message: "x" }), /^Sent to the "one" Teams workflow/);
    await service.webhooks.add("two", URL_B);
    const ambiguous = await tools.teams_chat_send_message.handler({ message: "x" });
    assert.match(ambiguous.textResultForLlm, /Several fixed-chat Teams workflows are saved \(one, two\)/);
    assert.match((await tools.teams_chat_send_message.handler({ webhook: "nope", message: "x" })).textResultForLlm, /No Teams workflow named "nope"/);
    assert.equal(webhook.calls.length, 1);
});

test("declining, or a client without confirmation UI, sends nothing", async (t) => {
    const { service, webhook, cleanup } = await setup();
    t.after(cleanup);
    await service.webhooks.add("chat", URL_A);
    const declined = await toolsFor(service, fakeSession({ confirm: false })).teams_chat_send_message.handler({ message: "x" });
    assert.equal(declined.resultType, "rejected");
    const noUi = await toolsFor(service, fakeSession({ elicitation: false })).teams_chat_send_message.handler({ message: "x" });
    assert.match(noUi.textResultForLlm, /teams-chat-compose/);
    assert.equal(webhook.calls.length, 0);
});

test("confirmBeforeSend: false sends without prompting", async (t) => {
    const { service, webhook, cleanup } = await setup({ config: { confirmBeforeSend: false } });
    t.after(cleanup);
    await service.webhooks.add("chat", URL_A);
    const session = fakeSession({ elicitation: false });
    assert.match(await toolsFor(service, session).teams_chat_send_message.handler({ message: "note", format: "text" }), /^Sent/);
    assert.equal(session.prompts.length, 0);
    assert.equal(webhook.calls.length, 1);
});

test("throttling is retried; server errors, gone workflows, and bad input are reported", async (t) => {
    const { service, webhook, cleanup } = await setup();
    t.after(cleanup);
    await service.webhooks.add("chat", URL_A);
    const tools = toolsFor(service, fakeSession());

    webhook.statuses.push(429);
    assert.match(await tools.teams_chat_send_message.handler({ message: "x" }), /^Sent/);
    assert.equal(webhook.calls.length, 2, "429 is retried once");

    webhook.statuses.push(502);
    assert.match((await tools.teams_chat_send_message.handler({ message: "x" })).textResultForLlm, /may still have been posted; check Teams/);
    assert.equal(webhook.calls.length, 3, "5xx is not retried");

    webhook.statuses.push(404);
    assert.match((await tools.teams_chat_send_message.handler({ message: "x" })).textResultForLlm, /workflow may be turned off, deleted/);

    assert.match((await tools.teams_chat_send_message.handler({ message: "  " })).textResultForLlm, /empty/);
    assert.match((await tools.teams_chat_send_message.handler({ message: "x".repeat(30_000) })).textResultForLlm, /too long/);
    assert.equal(webhook.calls.length, 4);
});

test("an approval that arrives after the confirmation window is ignored", async (t) => {
    const { service, webhook, cleanup } = await setup();
    t.after(cleanup);
    await service.webhooks.add("chat", URL_A);
    let clock = 0;
    const session = fakeSession({
        confirm: async () => {
            clock += 11 * 60_000;
            return true;
        },
    });
    const tools = Object.fromEntries(createTools({ service, session, now: () => clock }).map((tool) => [tool.name, tool]));
    const late = await tools.teams_chat_send_message.handler({ message: "x" });
    assert.equal(late.resultType, "rejected");
    assert.match(late.textResultForLlm, /more than 10 minutes after it was asked, so it was ignored/);
    assert.equal(webhook.calls.length, 0);
    const lateAdd = await tools.teams_chat_add_webhook.handler({ name: "other", url: URL_B });
    assert.equal(lateAdd.resultType, "rejected");
    assert.equal((await service.webhooks.list()).length, 1);
});

test("a webhook whose URL changes after confirmation is not used", async (t) => {
    const { service, webhook, cleanup } = await setup();
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
