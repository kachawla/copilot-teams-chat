import assert from "node:assert/strict";
import test from "node:test";
import { createTools } from "../com.github.copilot/extensions/teams-chat/lib/tools.mjs";
import { fakeSession, setup } from "./helpers/setup.mjs";

function toolsFor(service, session) {
    return Object.fromEntries(createTools({ service, session }).map((tool) => [tool.name, tool]));
}

test("tool names are namespaced and read-only tools skip the permission prompt", async (t) => {
    const { service, cleanup } = await setup({ signedIn: false });
    t.after(cleanup);
    const tools = createTools({ service, session: fakeSession() });
    for (const tool of tools) assert.match(tool.name, /^teams_chat_/);
    const skip = tools.filter((tool) => tool.skipPermission).map((tool) => tool.name).sort();
    assert.deepEqual(skip, ["teams_chat_find_people", "teams_chat_list_chats", "teams_chat_send_message", "teams_chat_status"]);
});

test("send asks for confirmation with the resolved recipient and message", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    const session = fakeSession({ confirm: true });
    const tools = toolsFor(service, session);
    const result = await tools.teams_chat_send_message.handler({ to: ["alice@contoso.com"], message: "Build is **green**" });
    assert.equal(typeof result, "string");
    assert.match(result, /^Sent to 1:1 chat with Alice Smith <alice@contoso.com>\./);
    assert.match(result, /Open in Teams: https:\/\/teams\.microsoft\.com/);
    assert.equal(session.prompts.length, 1);
    assert.match(session.prompts[0], /Send this Microsoft Teams message to 1:1 chat with Alice Smith <alice@contoso.com>\?/);
    assert.match(session.prompts[0], /Build is \*\*green\*\*/);
    assert.equal(cloud.state.messages[0].body.content, "<p>Build is <strong>green</strong></p>");
    assert.equal(session.logs[0].message, "Sent a Teams message to 1:1 chat with Alice Smith <alice@contoso.com>.");
});

test("declining the confirmation sends nothing and creates no chat", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession({ confirm: false }));
    const result = await tools.teams_chat_send_message.handler({ to: ["alice@contoso.com", "bob@contoso.com"], message: "hi" });
    assert.equal(result.resultType, "rejected");
    assert.equal(cloud.state.messages.length, 0);
    assert.equal(cloud.state.createdChats.length, 0);
});

test("without a confirmation UI, send refuses and points to the compose canvas", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession({ elicitation: false }));
    const result = await tools.teams_chat_send_message.handler({ to: ["me"], message: "hi" });
    assert.equal(result.resultType, "failure");
    assert.match(result.textResultForLlm, /teams-chat-compose/);
    assert.equal(cloud.state.messages.length, 0);
    assert.match(await tools.teams_chat_status.handler({}), /can't show confirmation prompts/);
});

test("confirmBeforeSend: false sends without prompting", async (t) => {
    const { service, cloud, cleanup } = await setup({ config: { confirmBeforeSend: false } });
    t.after(cleanup);
    const session = fakeSession({ elicitation: false });
    const result = await toolsFor(service, session).teams_chat_send_message.handler({ to: ["me"], message: "note", format: "text" });
    assert.match(result, /^Sent to your Notes to self chat/);
    assert.equal(session.prompts.length, 0);
    assert.equal(cloud.state.messages[0].body.content, "note");
});

test("ambiguous recipients come back to the model as candidates", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const result = await toolsFor(service, fakeSession()).teams_chat_send_message.handler({ to: ["Smith"], message: "hi" });
    assert.equal(result.resultType, "failure");
    assert.match(result.textResultForLlm, /matches more than one person/);
    assert.match(result.textResultForLlm, /Alice Smith <alice@contoso.com> \(PM\)/);
    assert.match(result.textResultForLlm, /Alex Smith <alex@contoso.com> \(Engineer\)/);
});

test("tools report not signed in, then sign-in returns instructions", async (t) => {
    const { service, cleanup } = await setup({ signedIn: false });
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession());
    assert.match(await tools.teams_chat_status.handler({}), /Not signed in/);
    const list = await tools.teams_chat_list_chats.handler({});
    assert.equal(list.resultType, "failure");
    assert.match(list.textResultForLlm, /teams_chat_sign_in/);

    const signIn = await tools.teams_chat_sign_in.handler({});
    assert.match(signIn, /Microsoft sign-in page opened/);
    assert.match(await tools.teams_chat_status.handler({}), /Sign-in in progress/);

    const device = await tools.teams_chat_sign_in.handler({ method: "device_code" });
    assert.match(device, /https:\/\/microsoft\.com\/devicelogin and enter the code ABCD-1234/);
});

test("sign-in when already signed in does not start a new flow", async (t) => {
    const { service, opened, cleanup } = await setup();
    t.after(cleanup);
    const result = await toolsFor(service, fakeSession()).teams_chat_sign_in.handler({});
    assert.match(result, /Already signed in to Microsoft Teams as Test User \(test@contoso.com\)/);
    assert.equal(opened.length, 1);
});

test("list chats filters by query and includes chat IDs", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession());
    const all = await tools.teams_chat_list_chats.handler({});
    assert.match(all, /1\. Alice Smith \(1:1, 2026-09-26 10:00 UTC\) chatId: 19:alice-chat@unq\.gbl\.spaces/);
    assert.match(all, /Notes to self chat has chatId 48:notes/);
    const filtered = await tools.teams_chat_list_chats.handler({ query: "carol" });
    assert.doesNotMatch(filtered, /Alice Smith \(1:1/);
    assert.match(filtered, /Bob Lee, Carol Diaz/);
    assert.match(filtered, /Launch plan/);
    assert.match(await tools.teams_chat_list_chats.handler({ query: "nothing-matches" }), /No recent chats match/);
});

test("find people by name or email", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession());
    assert.equal(await tools.teams_chat_find_people.handler({ query: "bob@contoso.com" }), "- Bob Lee <bob@contoso.com>");
    const smiths = await tools.teams_chat_find_people.handler({ query: "smith" });
    assert.equal(smiths.split("\n").length, 2);
});

test("sign out tool clears the session", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const tools = toolsFor(service, fakeSession());
    assert.match(await tools.teams_chat_sign_out.handler({}), /Signed out/);
    assert.match(await tools.teams_chat_status.handler({}), /Not signed in/);
});
