import assert from "node:assert/strict";
import fs from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createComposeCanvas } from "../com.github.copilot/extensions/teams-chat/lib/canvas.mjs";
import { fakeSession, linkParams, setup } from "./helpers/setup.mjs";

class FakeCanvasError extends Error {
    constructor(code, message) {
        super(message);
        this.code = code;
    }
}

async function openCanvas(t, { input = {}, handles } = {}) {
    const ctx = await setup({ handles });
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
    return { ...ctx, canvas, openedCanvas: opened, url, token, api, session, draftsDir };
}

test("canvas declaration exposes draft actions and nothing that sends", async (t) => {
    const { canvas } = await openCanvas(t);
    assert.equal(canvas.id, "teams-chat-compose");
    assert.deepEqual(canvas.actions.map((a) => a.name), ["update_draft", "get_draft"]);
});

test("open stores the draft and serves the page with a strict CSP", async (t) => {
    const { url, openedCanvas, api, draftsDir, service } = await openCanvas(t, { input: { to: ["alice@contoso.com"], message: "Hello" } });
    await service.contacts.save("Bob", "bob@contoso.com");
    assert.equal(openedCanvas.title, "Teams message");
    assert.equal(url.hostname, "127.0.0.1");
    const page = await fetch(url);
    assert.equal(page.status, 200);
    assert.match(page.headers.get("content-security-policy"), /script-src 'self'/);
    assert.match(await page.text(), /Nothing is sent until you press Send in Teams/);
    for (const asset of ["/app.js", "/style.css", "/teams-link.mjs"]) {
        assert.equal((await fetch(new URL(asset, url))).status, 200, asset);
    }
    const { json } = await api("/api/state");
    assert.deepEqual(json.draft.to, ["alice@contoso.com"]);
    assert.equal(json.draft.message, "Hello");
    assert.deepEqual(json.contacts, ["Bob"]);
    const saved = JSON.parse(await fs.readFile(path.join(draftsDir, "default.json"), "utf8"));
    assert.equal(saved.message, "Hello");
});

test("API requires the panel token and a loopback Host header", async (t) => {
    const { url, token, api } = await openCanvas(t);
    assert.equal((await fetch(new URL("/api/state", url))).status, 403);
    assert.equal((await api("/api/state", undefined, { "x-teams-chat-token": "wrong" })).status, 403);
    // Query token is accepted for EventSource GETs only, never for mutations.
    assert.equal((await fetch(new URL(`/api/open?token=${token}`, url), { method: "POST", body: "{}" })).status, 403);
    assert.equal((await api("/api/send", {})).status, 404);

    const status = await new Promise((resolve, reject) => {
        const req = http.request({ host: "127.0.0.1", port: url.port, path: "/", headers: { Host: "evil.example.com" } }, (res) =>
            resolve(res.statusCode),
        );
        req.on("error", reject);
        req.end();
    });
    assert.equal(status, 421);
});

test("Open in Teams opens the chat link and keeps the draft", async (t) => {
    const { api, opened, session } = await openCanvas(t, { input: { to: ["alice@contoso.com"], message: "- one\n- two" } });
    const res = await api("/api/open", {});
    assert.equal(res.status, 200);
    assert.equal(res.json.openedIn, "app");
    assert.equal(res.json.label, "your chat with alice@contoso.com");
    assert.match(res.json.link, /^https:\/\/teams\.microsoft\.com\/l\/chat\/0\/0\?/);
    assert.deepEqual(linkParams(opened[0]), { users: "alice@contoso.com", message: "• one\n• two" });
    assert.equal(res.json.state.draft.message, "- one\n- two", "the draft stays until the user clears it");
    assert.equal(res.json.state.draft.lastResult.openedIn, "app");
    assert.match(session.logs[0].message, /from the compose canvas/);
});

test("when nothing opens, the page still gets the link", async (t) => {
    const { api } = await openCanvas(t, { input: { to: ["a@x.com"], message: "hi" }, handles: () => false });
    const res = await api("/api/open", {});
    assert.equal(res.json.openedIn, null);
    assert.match(res.json.link, /^https:\/\/teams\.microsoft\.com/);
});

test("open errors are reported to the page", async (t) => {
    const { api, opened } = await openCanvas(t, { input: { to: ["Alice"], message: "hi" } });
    const res = await api("/api/open", {});
    assert.equal(res.status, 500);
    assert.equal(res.json.error, 'No email address is saved for "Alice". Use their email address, or ask Copilot to save them as a contact.');
    assert.equal(opened.length, 0);
});

test("agent actions update and read the same draft; edits push over SSE", async (t) => {
    const { canvas, url, token, api } = await openCanvas(t);
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
    await updateDraft.handler({ instanceId: "panel-1", input: { to: ["a@x.com"], message: "Agent draft" } });
    const pushed = await nextEvent();
    assert.equal(pushed.draft.message, "Agent draft");
    assert.deepEqual(pushed.draft.to, ["a@x.com"]);

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

test("invalid draft IDs are rejected on open", async (t) => {
    const { canvas } = await openCanvas(t);
    await assert.rejects(canvas.open({ instanceId: "panel-2", input: { draftId: "../escape" } }), /Invalid draftId/);
});
