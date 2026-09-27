import assert from "node:assert/strict";
import test from "node:test";
import { SELF_CHAT_ID, RecipientError, executeSend, planSend, renderMessage } from "../com.github.copilot/extensions/teams-chat/lib/send.mjs";
import { splitRecipients } from "../com.github.copilot/extensions/teams-chat/lib/recipients.mjs";
import { setup } from "./helpers/setup.mjs";

const account = { id: "me-id", displayName: "Test User", username: "test@contoso.com" };

test("one recipient by email resolves to the 1:1 chat", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    const plan = await planSend({ graph: service.graph, account, to: ["alice@contoso.com"] });
    assert.equal(plan.kind, "oneOnOne");
    assert.equal(plan.label, "1:1 chat with Alice Smith <alice@contoso.com>");
    assert.equal(cloud.state.createdChats.length, 0, "planning must not change anything");

    const result = await executeSend({ graph: service.graph, account, plan, html: "<p>hi</p>" });
    assert.equal(result.chatId, "19:alice-chat@unq.gbl.spaces");
    const [msg] = cloud.state.messages;
    assert.deepEqual(msg.body, { contentType: "html", content: "<p>hi</p>" });
    const create = cloud.state.requests.find((r) => r.method === "POST" && r.path === "/v1.0/chats");
    const body = JSON.parse(create.body);
    assert.equal(body.chatType, "oneOnOne");
    assert.deepEqual(
        body.members.map((m) => m["user@odata.bind"]),
        [`${cloud.graphBaseUrl}/users('me-id')`, `${cloud.graphBaseUrl}/users('alice-id')`],
    );
});

test("ambiguous names return candidates instead of guessing", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    await assert.rejects(
        planSend({ graph: service.graph, account, to: ["Smith"] }),
        (err) => err instanceof RecipientError && err.code === "ambiguous_recipient" && err.candidates.length === 2,
    );
});

test("an exact display name wins over partial matches", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const plan = await planSend({ graph: service.graph, account, to: ["alice smith"] });
    assert.equal(plan.people[0].id, "alice-id");
});

test("an email that only prefix-matches someone else is not accepted", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    await assert.rejects(planSend({ graph: service.graph, account, to: ["jane@contoso.com"] }), /No one in your organization has the address/);
    const plan = await planSend({ graph: service.graph, account, to: ["jane.doe@contoso.com"] });
    assert.equal(plan.people[0].id, "jane-id", "mail that differs from the UPN still resolves exactly");
});

test("directory names written Last, First stay one recipient", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const plan = await planSend({ graph: service.graph, account, to: ["Nguyen, Minh"] });
    assert.equal(plan.kind, "oneOnOne");
    assert.equal(plan.people[0].id, "minh-id");
});

test("splitRecipients", () => {
    assert.deepEqual(splitRecipients("a@x.com, b@x.com"), ["a@x.com", "b@x.com"]);
    assert.deepEqual(splitRecipients(["Smith, John"]), ["Smith, John"]);
    assert.deepEqual(splitRecipients("Smith, John; b@x.com\nme"), ["Smith, John", "b@x.com", "me"]);
    assert.deepEqual(splitRecipients(["a@x.com, Bob"]), ["a@x.com, Bob"]);
    assert.deepEqual(splitRecipients(undefined), []);
});

test("sends are not retried on 503/504, since they may have gone through", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    cloud.state.failNext["POST /v1.0/chats/48:notes/messages"] = [504];
    await assert.rejects(service.graph.sendChatMessage("48:notes", "<p>x</p>"), /may still have gone through; check Teams/);
    assert.equal(cloud.state.messages.length, 0);
    assert.equal(cloud.state.requests.filter((r) => r.method === "POST" && r.path.includes("/messages")).length, 1);

    cloud.state.failNext["POST /v1.0/chats/48:notes/messages"] = [429];
    await service.graph.sendChatMessage("48:notes", "<p>y</p>");
    assert.equal(cloud.state.messages.length, 1, "429 is retried");
});

test("unknown recipients fail clearly", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    await assert.rejects(planSend({ graph: service.graph, account, to: ["nobody@contoso.com"] }), /No one in your organization/);
    await assert.rejects(planSend({ graph: service.graph, account, to: ["Zed"] }), /Try their email address/);
    await assert.rejects(planSend({ graph: service.graph, account, to: [] }), /chatId or at least one recipient/);
});

test("me and the user's own address go to Notes to self", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    for (const to of [["me"], ["Notes to self"], ["test@contoso.com"]]) {
        const plan = await planSend({ graph: service.graph, account, to });
        assert.equal(plan.kind, "self");
        assert.equal(plan.chatId, SELF_CHAT_ID);
    }
    const plan = await planSend({ graph: service.graph, account, to: ["me"] });
    await executeSend({ graph: service.graph, account, plan, html: "<p>note</p>" });
    assert.equal(cloud.state.messages[0].chatId, "48:notes");
});

test("several recipients reuse the group chat with exactly those members", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    const plan = await planSend({ graph: service.graph, account, to: ["bob@contoso.com, carol@contoso.com"] });
    assert.equal(plan.kind, "existing");
    assert.equal(plan.chatId, "19:team-chat@thread.v2");
    await executeSend({ graph: service.graph, account, plan, html: "<p>x</p>" });
    assert.equal(cloud.state.createdChats.length, 0);
});

test("several recipients without a matching chat create a group chat", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    const plan = await planSend({ graph: service.graph, account, to: ["alice@contoso.com", "bob@contoso.com"], topic: "Sync" });
    assert.equal(plan.kind, "newGroup");
    assert.match(plan.label, /new group chat with Alice Smith <alice@contoso.com>, Bob Lee <bob@contoso.com> named "Sync"/);
    const result = await executeSend({ graph: service.graph, account, plan, html: "<p>x</p>" });
    assert.equal(cloud.state.createdChats.length, 1);
    assert.equal(cloud.state.createdChats[0].requestBody.topic, "Sync");
    assert.equal(result.chatId, cloud.state.createdChats[0].id);
});

test("chatId targets an existing chat and gets a readable label", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    const plan = await planSend({ graph: service.graph, account, chatId: "19:launch@thread.v2" });
    assert.equal(plan.label, 'chat "Launch plan"');
    const one = await planSend({ graph: service.graph, account, chatId: "19:alice-chat@unq.gbl.spaces" });
    assert.equal(one.label, "1:1 chat with Alice Smith");
    await assert.rejects(planSend({ graph: service.graph, account, chatId: "19:missing" }), /was not found/);
});

test("renderMessage rejects empty and oversized messages", () => {
    assert.throws(() => renderMessage("   "), /empty/);
    assert.throws(() => renderMessage("x".repeat(30_000)), /too long/);
    assert.equal(renderMessage("**hi**"), "<p><strong>hi</strong></p>");
});

test("graph retries throttled requests and refreshes after a 401", async (t) => {
    const { service, cloud, cleanup } = await setup();
    t.after(cleanup);
    cloud.state.failNext["GET /v1.0/me/chats"] = [429, 503];
    const chats = await service.graph.listChats({ myUserId: "me-id" });
    assert.equal(chats.length, 3);

    cloud.state.validAccessTokens.clear();
    const again = await service.graph.listChats({ myUserId: "me-id" });
    assert.equal(again.length, 3);
    assert.equal(cloud.state.tokenCounter, 2, "a new access token was fetched with the refresh token");
});

for (const status of [400, 403]) {
    test(`listChats falls back when $orderby is rejected with ${status}, and labels chats`, async (t) => {
        const { service, cloud, cleanup } = await setup();
        t.after(cleanup);
        cloud.state.rejectOrderBy = status;
        const chats = await service.graph.listChats({ myUserId: "me-id" });
        assert.deepEqual(
            chats.map((c) => c.label),
            ["Alice Smith", "Bob Lee, Carol Diaz", "Launch plan"],
        );
    });
}

test("graph refuses to send the token to non-Graph URLs", async (t) => {
    const { service, cleanup } = await setup();
    t.after(cleanup);
    await assert.rejects(service.graph.request("GET", "https://evil.example.com/v1.0/me"), /non-Graph URL/);
});
