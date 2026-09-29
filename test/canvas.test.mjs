import assert from "node:assert/strict";
import fs from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createComposeCanvas } from "../com.github.copilot/extensions/teams-chat/lib/canvas.mjs";
import { URL_A, URL_B, fakeSession, setup } from "./helpers/setup.mjs";

class FakeCanvasError extends Error {
    constructor(code, message) {
        super(message);
        this.code = code;
    }
}

async function openCanvas(t, { input = {}, webhooks = [["Team chat", URL_A]] } = {}) {
    const ctx = await setup();
    for (const [name, url] of webhooks) await ctx.service.webhooks.add(name, url);
    const draftsDir = await fs.mkdtemp(path.join(os.tmpdir(), "teams-chat-drafts-"));
    const session = fakeSession();
    const canvas = createComposeCanvas({
        createCanvas: (options) => options,
        CanvasError: FakeCanvasError,
        service: ctx.service,
        session,
        draftsDir: () => draftsDir,
    });
    const opened = await canvas.open({ instanceId: "panel-1", canvasId: canvas.id, input });
    const url = new URL(opened.url);
    const token = url.searchParams.get("token");
    const api = async (p, body, headers = {}) => {
        const res = await fetch(new URL(p, url), {
            method: body === undefined ? "GET" : "POST",
            headers: { "x-teams-chat-token": token, "Content-Type": "application/json", ...headers },
            body: body === undefined ? undefined : JSON.stringify(body),
        });
        return { status: res.status, json: await res.json().catch(() => null) };
    };
    t.after(async () => {
        await canvas.onClose({ instanceId: "panel-1" });
        await ctx.cleanup();
        await fs.rm(draftsDir, { recursive: true, force: true });
    });
    return { ...ctx, canvas, opened, url, token, api, session, draftsDir };
}

test("canvas declaration exposes draft actions but no agent send action", async (t) => {
    const { canvas } = await openCanvas(t);
    assert.equal(canvas.id, "teams-chat-compose");
    assert.deepEqual(canvas.actions.map((a) => a.name), ["update_draft", "get_draft"]);
});

test("open stores the draft and serves the page with a strict CSP", async (t) => {
    const { url, opened, api, draftsDir } = await openCanvas(t, { input: { webhook: "Team chat", message: "Hello" } });
    assert.equal(opened.title, "Teams message");
    assert.equal(url.hostname, "127.0.0.1");
    const page = await fetch(url);
    assert.equal(page.status, 200);
    assert.match(page.headers.get("content-security-policy"), /script-src 'self'/);
    assert.match(await page.text(), /<title>Teams message<\/title>/);
    for (const asset of ["/app.js", "/style.css", "/markdown.mjs"]) {
        assert.equal((await fetch(new URL(asset, url))).status, 200, asset);
    }
    const { json } = await api("/api/state");
    assert.equal(json.draft.webhook, "Team chat");
    assert.equal(json.draft.message, "Hello");
    assert.deepEqual(json.webhooks.map((w) => w.name), ["Team chat"]);
    assert.equal(JSON.stringify(json).includes("sig=secret"), false, "webhook URLs are not exposed to the page");
    const saved = JSON.parse(await fs.readFile(path.join(draftsDir, "default.json"), "utf8"));
    assert.equal(saved.message, "Hello");
});

test("API requires the panel token and a loopback Host header", async (t) => {
    const { url, token, api } = await openCanvas(t);
    assert.equal((await fetch(new URL("/api/state", url))).status, 403);
    assert.equal((await api("/api/state", undefined, { "x-teams-chat-token": "wrong" })).status, 403);
    // Query token is accepted for EventSource GETs only, never for mutations.
    const post = await fetch(new URL(`/api/draft?token=${token}`, url), { method: "POST", body: "{}" });
    assert.equal(post.status, 403);
    assert.equal((await api("/api/sign-in", {})).status, 404);

    const status = await new Promise((resolve, reject) => {
        const req = http.request({ host: "127.0.0.1", port: url.port, path: "/", headers: { Host: "evil.example.com" } }, (res) =>
            resolve(res.statusCode),
        );
        req.on("error", reject);
        req.end();
    });
    assert.equal(status, 421);
});

test("user send flow: confirm step, then send, then message is cleared", async (t) => {
    const { api, webhook, session } = await openCanvas(t, { input: { message: "- one\n- two" } });

    const first = await api("/api/send", { confirm: false });
    assert.equal(first.json.needsConfirm, true);
    assert.equal(first.json.label, 'the "Team chat" Teams workflow', "the only saved workflow is used");
    assert.match(first.json.fingerprint, /^[A-Za-z0-9_-]{43}$/);
    assert.equal(webhook.calls.length, 0);

    const stale = await api("/api/send", { confirm: true, fingerprint: "stale" });
    assert.equal(stale.json.needsConfirm, true, "a stale confirmation does not send");
    assert.equal(webhook.calls.length, 0);

    const sent = await api("/api/send", { confirm: true, fingerprint: first.json.fingerprint });
    assert.equal(sent.json.sent, true);
    assert.equal(webhook.calls.length, 1);
    assert.equal(webhook.calls[0].body.attachments[0].content.body[0].text, "- one\n- two");
    const draft = sent.json.state.draft;
    assert.equal(draft.message, "");
    assert.equal(draft.webhook, "Team chat");
    assert.equal(draft.lastResult.ok, true);
    assert.equal(draft.lastResult.messagePreview, "- one\n- two");
    assert.match(session.logs[0].message, /from the compose canvas/);
});

test("a draft changed after the user confirmed is not sent", async (t) => {
    const { api, canvas, webhook } = await openCanvas(t, {
        input: { webhook: "Team chat", message: "approved text" },
        webhooks: [["Team chat", URL_A], ["Release", URL_B]],
    });
    const first = await api("/api/send", { confirm: false });
    const updateDraft = canvas.actions.find((a) => a.name === "update_draft");

    await updateDraft.handler({ instanceId: "panel-1", input: { message: "different text" } });
    assert.equal((await api("/api/send", { confirm: true, fingerprint: first.json.fingerprint })).json.needsConfirm, true);

    await updateDraft.handler({ instanceId: "panel-1", input: { message: "approved text", webhook: "Release" } });
    const changed = await api("/api/send", { confirm: true, fingerprint: first.json.fingerprint });
    assert.equal(changed.json.needsConfirm, true);
    assert.equal(webhook.calls.length, 0);

    const sent = await api("/api/send", { confirm: true, fingerprint: changed.json.fingerprint });
    assert.equal(sent.json.sent, true);
    assert.equal(webhook.calls[0].url, URL_B);
});

test("send errors are reported to the page", async (t) => {
    const { api } = await openCanvas(t, { input: { message: "hi" }, webhooks: [["a", URL_A], ["b", URL_B]] });
    const res = await api("/api/send", { confirm: false });
    assert.equal(res.status, 500);
    assert.match(res.json.error, /Several fixed-chat Teams workflows are saved/);
});

test("agent actions update and read the same draft; edits and new webhooks push over SSE", async (t) => {
    const { canvas, url, token, api, service } = await openCanvas(t);
    const events = await fetch(new URL(`/api/events?token=${token}`, url));
    const reader = events.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    const nextEvent = async () => {
        while (!buffer.includes("\n\n")) buffer += decoder.decode((await reader.read()).value);
        const [frame, ...rest] = buffer.split("\n\n");
        buffer = rest.join("\n\n");
        return JSON.parse(frame.replace(/^data: /, ""));
    };
    assert.equal((await nextEvent()).draft.message, "");

    const updateDraft = canvas.actions.find((a) => a.name === "update_draft");
    await updateDraft.handler({ instanceId: "panel-1", input: { message: "Agent draft" } });
    assert.equal((await nextEvent()).draft.message, "Agent draft");

    await service.webhooks.add("Release", URL_B);
    assert.deepEqual((await nextEvent()).webhooks.map((w) => w.name), ["Release", "Team chat"]);

    await api("/api/draft", { message: "User edit" });
    const draft = await canvas.actions.find((a) => a.name === "get_draft").handler({ instanceId: "panel-1" });
    assert.equal(draft.message, "User edit");
    await reader.cancel();
});

test("replaying the same open input does not overwrite later edits", async (t) => {
    const { canvas, api } = await openCanvas(t, { input: { message: "first" } });
    await api("/api/draft", { message: "edited by the user" });
    await canvas.open({ instanceId: "panel-1", input: { message: "first" } });
    assert.equal((await api("/api/state")).json.draft.message, "edited by the user");
    await canvas.open({ instanceId: "panel-1", input: { message: "new request" } });
    assert.equal((await api("/api/state")).json.draft.message, "new request");
});

test("the canvas sends to people through a send-to-anyone workflow", async (t) => {
    const { api, service, webhook } = await openCanvas(t, { input: { to: ["Alice"], message: "hello" } });
    await service.webhooks.add("direct", "https://prod-3.westus.logic.azure.com/workflows/d?sig=d", null, "direct");
    await service.contacts.save("Alice", "alice@contoso.com");
    const state = (await api("/api/state")).json;
    assert.deepEqual(state.contacts, ["Alice"]);
    assert.deepEqual(state.webhooks.map((w) => [w.name, w.kind]), [["direct", "direct"], ["Team chat", "chat"]]);

    const first = await api("/api/send", { confirm: false });
    assert.equal(first.json.label, "Alice <alice@contoso.com> (1:1 chat, as you)");
    await service.contacts.save("Alice", "alice2@contoso.com");
    const changed = await api("/api/send", { confirm: true, fingerprint: first.json.fingerprint });
    assert.equal(changed.json.needsConfirm, true, "a contact that now points elsewhere needs a new confirmation");
    const sent = await api("/api/send", { confirm: true, fingerprint: changed.json.fingerprint });
    assert.equal(sent.json.sent, true);
    assert.equal(webhook.calls[0].body.to, "alice2@contoso.com");
});

test("invalid draft IDs are rejected on open", async (t) => {
    const { canvas } = await openCanvas(t);
    await assert.rejects(canvas.open({ instanceId: "panel-2", input: { draftId: "../escape" } }), /Invalid draftId/);
});
