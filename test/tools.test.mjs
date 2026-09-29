import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { normalizeRecipients } from "../com.github.copilot/extensions/teams-chat/lib/compose.mjs";
import { createTools } from "../com.github.copilot/extensions/teams-chat/lib/tools.mjs";
import { fakeSession, linkParams, setup } from "./helpers/setup.mjs";

function toolsFor(service, session = fakeSession()) {
    return Object.fromEntries(createTools({ service, session }).map((tool) => [tool.name, tool]));
}

test("tools are namespaced; only read-only and draft tools skip the permission prompt", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const tools = createTools({ service, session: fakeSession() });
    assert.deepEqual(tools.map((tool) => tool.name), [
        "teams_chat_draft_message",
        "teams_chat_save_contact",
        "teams_chat_list_contacts",
        "teams_chat_remove_contact",
    ]);
    assert.deepEqual(tools.filter((tool) => tool.skipPermission).map((tool) => tool.name).sort(), ["teams_chat_draft_message", "teams_chat_list_contacts"]);
});

test("drafting to one person opens the Teams app with the plain-text message and says it isn't sent", async (t) => {
    const { service, opened, cleanup } = await setup();
    t.after(cleanup);
    const session = fakeSession();
    const result = await toolsFor(service, session).teams_chat_draft_message.handler({ to: ["alice@contoso.com"], message: "Build is **green**" });
    assert.equal(opened.length, 1);
    assert.match(opened[0], /^msteams:\/l\/chat\/0\/0\?/);
    assert.deepEqual(linkParams(opened[0]), { users: "alice@contoso.com", message: "Build is green" });
    assert.match(result, /^Opened your chat with alice@contoso\.com in Teams with the message ready\. It is NOT sent yet/);
    assert.match(result, /https:\/\/teams\.microsoft\.com\/l\/chat\/0\/0\?users=alice%40contoso\.com/);
    assert.equal(session.logs[0].message, "Opened a Teams draft to your chat with alice@contoso.com.");
});

test("falls back to Teams on the web when the app isn't installed, and to a link when nothing opens", async (t) => {
    const noApp = await setup({ handles: (url) => url.startsWith("https:") });
    t.after(noApp.cleanup);
    const web = await toolsFor(noApp.service).teams_chat_draft_message.handler({ to: ["a@x.com"], message: "hi" });
    assert.equal(noApp.opened.length, 2);
    assert.match(noApp.opened[1], /^https:\/\/teams\.microsoft\.com\//);
    assert.match(web, /in Teams on the web/);

    const nothing = await setup({ handles: () => false });
    t.after(nothing.cleanup);
    const link = await toolsFor(nothing.service).teams_chat_draft_message.handler({ to: ["a@x.com"], message: "hi" });
    assert.match(link, /^Couldn't open Teams automatically\. Give the user this link .*https:\/\/teams\.microsoft\.com/);
});

test("openIn: web and tenantId come from config", async (t) => {
    const { service, opened, cleanup } = await setup({ config: { openIn: "web", tenantId: "contoso.onmicrosoft.com" } });
    t.after(cleanup);
    await toolsFor(service).teams_chat_draft_message.handler({ to: ["a@x.com"], message: "hi" });
    assert.equal(opened.length, 1);
    assert.equal(linkParams(opened[0]).tenantId, "contoso.onmicrosoft.com");
    assert.match(opened[0], /^https:/);
});

test("several people open a group chat with a topic", async (t) => {
    const { service, opened, cleanup } = await setup();
    t.after(cleanup);
    const result = await toolsFor(service).teams_chat_draft_message.handler({
        to: ["a@x.com, b@x.com"],
        topic: "Release sync",
        message: "hi",
    });
    assert.deepEqual(linkParams(opened[0]), { users: "a@x.com,b@x.com", topicName: "Release sync", message: "hi" });
    assert.match(result, /Opened a group chat with a@x\.com, b@x\.com \(named "Release sync" if it's new\)/);
});

test("contact names resolve to saved addresses; unknown names are never guessed", async (t) => {
    const { service, opened, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service);
    const unknown = await tools.teams_chat_draft_message.handler({ to: ["Alice"], message: "hi" });
    assert.equal(unknown.resultType, "failure");
    assert.match(unknown.textResultForLlm, /No email address is known for "Alice"\. Ask the user for the address \(don't guess it\)/);
    assert.equal(opened.length, 0);

    assert.equal(await tools.teams_chat_save_contact.handler({ name: "Alice Smith", email: "alice@contoso.com" }), "Saved Alice Smith <alice@contoso.com>.");
    assert.equal(await tools.teams_chat_list_contacts.handler({}), "- Alice Smith <alice@contoso.com>");
    const result = await tools.teams_chat_draft_message.handler({ to: ["alice smith", "ALICE@contoso.com"], message: "hi" });
    assert.match(result, /your chat with Alice Smith <alice@contoso\.com>/, "duplicates collapse to one person");
    assert.equal(linkParams(opened[0]).users, "alice@contoso.com");

    assert.match(await tools.teams_chat_save_contact.handler({ name: "alice smith", email: "asmith@contoso.com" }), /alice@contoso\.com is now asmith@contoso\.com/);
    assert.match(await tools.teams_chat_remove_contact.handler({ name: "Alice Smith" }), /Removed/);
    assert.equal(await tools.teams_chat_list_contacts.handler({}), "No contacts saved.");
    assert.equal((await tools.teams_chat_remove_contact.handler({ name: "gone" })).resultType, "failure");
});

test("contacts are validated and stored privately", async (t) => {
    const { service, dir, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service);
    assert.match((await tools.teams_chat_save_contact.handler({ name: "Bob", email: "not-an-email" })).textResultForLlm, /not a valid email/);
    assert.match((await tools.teams_chat_save_contact.handler({ name: "a@b.com", email: "a@b.com" })).textResultForLlm, /can't be an email/);
    assert.match((await tools.teams_chat_save_contact.handler({ name: "Bob", email: "Bob <bob@x.com>" })).textResultForLlm, /not a valid email/);
    assert.match(await tools.teams_chat_save_contact.handler({ name: "José Núñez", email: "jose@contoso.com" }), /Saved/);
    if (process.platform !== "win32") assert.equal((await fs.stat(path.join(dir, "contacts.json"))).mode & 0o777, 0o600);
});

test("bad input fails without opening Teams", async (t) => {
    const { service, opened, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service);
    assert.match((await tools.teams_chat_draft_message.handler({ to: [], message: "hi" })).textResultForLlm, /Say who the message is for/);
    assert.match((await tools.teams_chat_draft_message.handler({ to: ["a@x.com"], message: "  " })).textResultForLlm, /empty/);
    assert.match((await tools.teams_chat_draft_message.handler({ to: ["a@x.com"], message: "x".repeat(9000) })).textResultForLlm, /too long/);
    assert.equal(opened.length, 0);
});

test("recipient parsing keeps Last, First names whole", () => {
    assert.deepEqual(normalizeRecipients("a@x.com, b@x.com; Smith, John\nme"), ["a@x.com", "b@x.com", "Smith, John", "me"]);
    assert.deepEqual(normalizeRecipients(["a@x.com, Bob"]), ["a@x.com, Bob"]);
    assert.deepEqual(normalizeRecipients(undefined), []);
});

test("config rejects a bad tenantId", async () => {
    const { loadConfig } = await import("../com.github.copilot/extensions/teams-chat/lib/config.mjs");
    const os = await import("node:os");
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "teams-chat-cfg-"));
    await fs.writeFile(path.join(dir, "config.json"), JSON.stringify({ tenantId: "bad/tenant" }));
    await assert.rejects(loadConfig({ dir }), /Invalid tenantId/);
    await fs.rm(dir, { recursive: true, force: true });
});
