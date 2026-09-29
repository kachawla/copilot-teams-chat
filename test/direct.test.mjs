import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { normalizeRecipients } from "../com.github.copilot/extensions/teams-chat/lib/delivery.mjs";
import { createTools } from "../com.github.copilot/extensions/teams-chat/lib/tools.mjs";
import { URL_A, URL_B, fakeSession, setup } from "./helpers/setup.mjs";

const DIRECT = "https://prod-3.westus.logic.azure.com/workflows/direct/triggers/manual/paths/invoke?sig=direct";

function toolsFor(service, session = fakeSession()) {
    return Object.fromEntries(createTools({ service, session }).map((tool) => [tool.name, tool]));
}

async function withDirect(t, extra = []) {
    const ctx = await setup();
    t.after(ctx.cleanup);
    await ctx.service.webhooks.add("direct", DIRECT, null, "direct");
    for (const [name, url] of extra) await ctx.service.webhooks.add(name, url);
    return ctx;
}

test("adding a send-to-anyone workflow says what it allows", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const session = fakeSession();
    const result = await toolsFor(service, session).teams_chat_add_webhook.handler({ name: "direct", url: DIRECT, kind: "direct" });
    assert.match(result, /Saved the send-to-anyone workflow "direct"/);
    assert.match(session.prompts[0], /as a send-to-anyone workflow\? Copilot will be able to send Teams messages as you/);
    assert.deepEqual((await service.webhooks.list()).map((w) => [w.name, w.kind]), [["direct", "direct"]]);
});

test("sending to one person posts the recipient and HTML to the direct workflow", async (t) => {
    const { service, webhook } = await withDirect(t, [["team chat", URL_A]]);
    const session = fakeSession();
    const result = await toolsFor(service, session).teams_chat_send_message.handler({
        to: ["alice@contoso.com"],
        message: "Build is **green**",
    });
    assert.match(result, /^Handed the message for alice@contoso\.com \(1:1 chat, as you\) to the "direct" workflow\./);
    assert.match(session.prompts[0], /^Send this Microsoft Teams message to alice@contoso\.com \(1:1 chat, as you\)\?\n\nBuild is \*\*green\*\*$/);
    assert.equal(webhook.calls.length, 1);
    assert.equal(webhook.calls[0].url, DIRECT);
    assert.deepEqual(webhook.calls[0].body, {
        to: "alice@contoso.com",
        topic: null,
        html: "<p>Build is <strong>green</strong></p>",
        text: "Build is **green**",
    });
});

test("several people make a named group chat", async (t) => {
    const { service, webhook } = await withDirect(t);
    const session = fakeSession();
    await toolsFor(service, session).teams_chat_send_message.handler({
        to: ["alice@contoso.com, bob@contoso.com"],
        topic: "Release sync",
        message: "hi",
    });
    assert.match(session.prompts[0], /a new group chat with alice@contoso\.com, bob@contoso\.com named "Release sync" \(as you\)/);
    assert.equal(webhook.calls[0].body.to, "alice@contoso.com;bob@contoso.com");
    assert.equal(webhook.calls[0].body.topic, "Release sync");
});

test("contact names resolve to saved addresses; unknown names are never guessed", async (t) => {
    const { service, webhook } = await withDirect(t);
    const tools = toolsFor(service);
    const unknown = await tools.teams_chat_send_message.handler({ to: ["Alice"], message: "hi" });
    assert.equal(unknown.resultType, "failure");
    assert.match(unknown.textResultForLlm, /No email address is known for "Alice"\. Ask the user for the address \(don't guess it\)/);
    assert.equal(webhook.calls.length, 0);

    assert.equal(await tools.teams_chat_save_contact.handler({ name: "Alice Smith", email: "alice@contoso.com" }), "Saved Alice Smith <alice@contoso.com>.");
    assert.equal(await tools.teams_chat_list_contacts.handler({}), "- Alice Smith <alice@contoso.com>");

    const session = fakeSession();
    await toolsFor(service, session).teams_chat_send_message.handler({ to: ["alice smith"], message: "hi" });
    assert.match(session.prompts[0], /to Alice Smith <alice@contoso\.com> \(1:1 chat, as you\)/);
    assert.equal(webhook.calls[0].body.to, "alice@contoso.com");

    assert.match(await tools.teams_chat_save_contact.handler({ name: "alice smith", email: "asmith@contoso.com" }), /alice@contoso\.com is now asmith@contoso\.com/);
    assert.match(await tools.teams_chat_remove_contact.handler({ name: "Alice Smith" }), /Removed/);
    assert.equal(await tools.teams_chat_list_contacts.handler({}), "No contacts saved.");
});

test("contacts are validated and stored privately", async (t) => {
    const { service, dir } = await withDirect(t);
    const tools = toolsFor(service);
    assert.match((await tools.teams_chat_save_contact.handler({ name: "Bob", email: "not-an-email" })).textResultForLlm, /not a valid email/);
    assert.match((await tools.teams_chat_save_contact.handler({ name: "a@b.com", email: "a@b.com" })).textResultForLlm, /can't be an email/);
    assert.match((await tools.teams_chat_save_contact.handler({ name: "Bob", email: "Bob <bob@x.com>" })).textResultForLlm, /not a valid email/);
    await tools.teams_chat_save_contact.handler({ name: "José Núñez", email: "jose@contoso.com" });
    const file = path.join(dir, "contacts.json");
    if (process.platform !== "win32") assert.equal((await fs.stat(file)).mode & 0o777, 0o600);
});

test("to needs a send-to-anyone workflow, and a fixed chat doesn't take to", async (t) => {
    const { service, cleanup, webhook } = await setup();
    t.after(cleanup);
    await service.webhooks.add("team chat", URL_A);
    const tools = toolsFor(service);
    const none = await tools.teams_chat_send_message.handler({ to: ["a@contoso.com"], message: "hi" });
    assert.match(none.textResultForLlm, /No send-to-anyone workflow is saved/);
    const wrong = await tools.teams_chat_send_message.handler({ webhook: "team chat", to: ["a@contoso.com"], message: "hi" });
    assert.match(wrong.textResultForLlm, /posts to one fixed chat, so it can't send to specific people/);

    await service.webhooks.add("direct", DIRECT, null, "direct");
    const noTo = await tools.teams_chat_send_message.handler({ webhook: "direct", message: "hi" });
    assert.match(noTo.textResultForLlm, /is a send-to-anyone workflow, so it needs to/);
    assert.match(await tools.teams_chat_send_message.handler({ message: "hi" }), /^Sent to the "team chat" Teams workflow/, "without to, the fixed chat is used");
    assert.equal(webhook.calls[0].url, URL_A);
});

test("recipient limits, duplicates, and parsing", async (t) => {
    const { service, webhook } = await withDirect(t);
    const tools = toolsFor(service);
    const many = Array.from({ length: 20 }, (_, i) => `p${i}@contoso.com`);
    assert.match((await tools.teams_chat_send_message.handler({ to: many, message: "hi" })).textResultForLlm, /at most 19 other people/);
    await tools.teams_chat_send_message.handler({ to: ["a@contoso.com", "A@contoso.com"], message: "hi" });
    assert.equal(webhook.calls[0].body.to, "a@contoso.com", "duplicates are removed, so it stays a 1:1 chat");
    assert.deepEqual(normalizeRecipients("a@x.com, b@x.com; Smith, John\nme"), ["a@x.com", "b@x.com", "Smith, John", "me"]);
});

test("a changed recipient after confirmation is not sent", async (t) => {
    const { service, webhook } = await withDirect(t, [["other", URL_B]]);
    await service.contacts.save("Alice", "alice@contoso.com");
    const session = fakeSession({
        confirm: async () => {
            await service.webhooks.add("direct", URL_B, null, "chat");
            return true;
        },
    });
    const result = await toolsFor(service, session).teams_chat_send_message.handler({ to: ["Alice"], message: "hi" });
    assert.match(result.textResultForLlm, /changed or was removed. Nothing was sent/);
    assert.equal(webhook.calls.length, 0);
});
