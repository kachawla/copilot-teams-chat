import assert from "node:assert/strict";
import fs from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createComposeCanvas } from "../com.github.copilot/extensions/teams-chat/lib/canvas.mjs";
import { fakeSession, setup } from "./helpers/setup.mjs";

class FakeCanvasError extends Error {
    constructor(code, message) {
        super(message);
        this.code = code;
    }
}

async function openCanvas(t, { signedIn = true, input = {} } = {}) {
    const ctx = await setup({ signedIn });
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
    const { url, opened, api, draftsDir } = await openCanvas(t, { input: { to: ["alice@contoso.com"], message: "Hello" } });
    assert.equal(opened.title, "Teams message");
    assert.equal(url.hostname, "127.0.0.1");
    const page = await fetch(url);
    assert.equal(page.status, 200);
    assert.match(page.headers.get("content-security-policy"), /script-src 'self'/);
    assert.match(await page.text(), /<title>Teams message<\/title>/);
    for (const asset of ["/app.js", "/style.css", "/markdown.mjs", "/recipients.mjs"]) {
        assert.equal((await fetch(new URL(asset, url))).status, 200, asset);
    }
    const { json } = await api("/api/state");
    assert.deepEqual(json.draft.to, ["alice@contoso.com"]);
    assert.equal(json.draft.message, "Hello");
    assert.equal(json.auth.signedIn, true);
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

    const status = await new Promise((resolve, reject) => {
        const req = http.request(
            { host: "127.0.0.1", port: url.port, path: "/", headers: { Host: "evil.example.com" } },
            (res) => resolve(res.statusCode),
        );
        req.on("error", reject);
        req.end();
    });
    assert.equal(status, 421);
});

test("user send flow: confirm step, then send, then draft is cleared", async (t) => {
    const { api, cloud, session } = await openCanvas(t, { input: { to: ["alice@contoso.com"], message: "- one\n- two" } });

    const first = await api("/api/send", { confirm: false });
    assert.equal(first.json.needsConfirm, true);
    assert.equal(first.json.label, "1:1 chat with Alice Smith <alice@contoso.com>");
    assert.match(first.json.fingerprint, /^[A-Za-z0-9_-]{43}$/);
    assert.equal(cloud.state.messages.length, 0);

    const mismatch = await api("/api/send", { confirm: true, fingerprint: "stale" });
    assert.equal(mismatch.json.needsConfirm, true, "a stale confirmation does not send");
    assert.equal(cloud.state.messages.length, 0);

    const sent = await api("/api/send", { confirm: true, fingerprint: first.json.fingerprint });
    assert.equal(sent.json.sent, true);
    assert.equal(cloud.state.messages.length, 1);
    assert.equal(cloud.state.messages[0].body.content, "<ul><li>one</li><li>two</li></ul>");
    const draft = sent.json.state.draft;
    assert.equal(draft.message, "");
    assert.equal(draft.chatId, "19:alice-chat@unq.gbl.spaces");
    assert.equal(draft.lastResult.ok, true);
    assert.equal(draft.lastResult.messagePreview, "- one\n- two");
    assert.match(session.logs[0].message, /from the compose canvas/);
});

test("a draft changed after the user confirmed is not sent", async (t) => {
    const { api, canvas, cloud } = await openCanvas(t, { input: { chatId: "19:launch@thread.v2", chatLabel: "Launch plan", message: "approved text" } });
    const first = await api("/api/send", { confirm: false });
    const updateDraft = canvas.actions.find((a) => a.name === "update_draft");

    await updateDraft.handler({ instanceId: "panel-1", input: { message: "different text" } });
    const changedMessage = await api("/api/send", { confirm: true, fingerprint: first.json.fingerprint });
    assert.equal(changedMessage.json.needsConfirm, true);

    await updateDraft.handler({ instanceId: "panel-1", input: { message: "approved text", chatId: "19:team-chat@thread.v2" } });
    const changedChat = await api("/api/send", { confirm: true, fingerprint: first.json.fingerprint });
    assert.equal(changedChat.json.needsConfirm, true);
    assert.equal(cloud.state.messages.length, 0);

    const sent = await api("/api/send", { confirm: true, fingerprint: changedChat.json.fingerprint });
    assert.equal(sent.json.sent, true);
    assert.equal(cloud.state.messages[0].chatId, "19:team-chat@thread.v2");
});

test("send errors are reported to the page", async (t) => {
    const { api } = await openCanvas(t, { input: { to: ["Smith"], message: "hi" } });
    const res = await api("/api/send", { confirm: false });
    assert.equal(res.status, 500);
    assert.match(res.json.error, /matches more than one person/);
    assert.match(res.json.error, /alex@contoso.com/);
});

test("agent actions update and read the same draft; edits push over SSE", async (t) => {
    const { canvas, url, token, api } = await openCanvas(t);
    const events = await fetch(new URL(`/api/events?token=${token}`, url));
    const reader = events.body.getReader();
    const decoder = new TextDecoder();
    const nextEvent = async () => {
        let buffer = "";
        while (!buffer.includes("\n\n")) buffer += decoder.decode((await reader.read()).value);
        return JSON.parse(buffer.split("\n\n")[0].replace(/^data: /, ""));
    };
    const initial = await nextEvent();
    assert.equal(initial.draft.message, "");

    const updateDraft = canvas.actions.find((a) => a.name === "update_draft");
    await updateDraft.handler({ instanceId: "panel-1", input: { chatId: "19:launch@thread.v2", chatLabel: "Launch plan", message: "Agent draft" } });
    const pushed = await nextEvent();
    assert.equal(pushed.draft.message, "Agent draft");
    assert.equal(pushed.draft.chatId, "19:launch@thread.v2");

    await api("/api/draft", { message: "User edit" });
    const getDraft = canvas.actions.find((a) => a.name === "get_draft");
    const draft = await getDraft.handler({ instanceId: "panel-1" });
    assert.equal(draft.message, "User edit");
    await reader.cancel();
});

test("choosing recipients clears a previously selected chat", async (t) => {
    const { api } = await openCanvas(t, { input: { chatId: "19:launch@thread.v2", chatLabel: "Launch plan" } });
    const { json } = await api("/api/draft", { to: "bob@contoso.com, carol@contoso.com; Nguyen, Minh" });
    assert.equal(json.draft.chatId, null);
    assert.deepEqual(json.draft.to, ["bob@contoso.com", "carol@contoso.com", "Nguyen, Minh"]);
});

test("sign-in from the canvas starts a flow and reports it", async (t) => {
    const { api } = await openCanvas(t, { signedIn: false });
    const { json } = await api("/api/sign-in", { method: "device_code" });
    assert.equal(json.auth.signedIn, false);
    assert.equal(json.auth.pending.userCode, "ABCD-1234");
});

test("replaying the same open input does not overwrite later edits", async (t) => {
    const { canvas, api } = await openCanvas(t, { input: { to: ["me"], message: "first" } });
    await api("/api/draft", { message: "edited by the user" });
    await canvas.open({ instanceId: "panel-1", input: { to: ["me"], message: "first" } });
    assert.equal((await api("/api/state")).json.draft.message, "edited by the user");
    await canvas.open({ instanceId: "panel-1", input: { to: ["me"], message: "new request" } });
    assert.equal((await api("/api/state")).json.draft.message, "new request");
});

test("invalid draft IDs are rejected on open", async (t) => {
    const { canvas } = await openCanvas(t);
    await assert.rejects(canvas.open({ instanceId: "panel-2", input: { draftId: "../escape" } }), /Invalid draftId/);
});
