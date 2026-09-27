import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { AuthManager, NotSignedInError } from "../com.github.copilot/extensions/teams-chat/lib/auth.mjs";
import { loadConfig } from "../com.github.copilot/extensions/teams-chat/lib/config.mjs";
import { TokenStore } from "../com.github.copilot/extensions/teams-chat/lib/token-store.mjs";
import { fakeBrowser } from "./helpers/mock-cloud.mjs";
import { setup } from "./helpers/setup.mjs";

test("browser sign-in uses PKCE and persists the refresh token privately", async (t) => {
    const ctx = await setup({ signedIn: false });
    t.after(ctx.cleanup);

    const pending = await ctx.service.auth.startSignIn({ method: "browser" });
    assert.equal(pending.method, "browser");
    assert.equal(ctx.opened.length, 1);
    const authorize = new URL(pending.url);
    assert.equal(authorize.searchParams.get("code_challenge_method"), "S256");
    assert.match(authorize.searchParams.get("redirect_uri"), /^http:\/\/localhost:\d+$/);
    assert.match(authorize.searchParams.get("scope"), /ChatMessage\.Send/);
    assert.match(authorize.searchParams.get("scope"), /offline_access/);

    const callback = await fakeBrowser(pending.url);
    assert.equal(callback.status, 200);
    assert.match(callback.body, /Signed in to Microsoft Teams/);

    const token = await ctx.service.auth.getAccessToken({ waitForPendingMs: 5000 });
    assert.equal(token, "at-1");
    const status = await ctx.service.auth.status();
    assert.equal(status.signedIn, true);
    assert.equal(status.pending, null);
    assert.deepEqual(status.account, { id: "me-id", displayName: "Test User", username: "test@contoso.com", tenantId: "tenant-id" });

    const file = path.join(ctx.dir, "auth.json");
    const saved = JSON.parse(await fs.readFile(file, "utf8"));
    assert.equal(saved.refreshToken, "rt-1");
    assert.equal(saved.accessToken, undefined, "access token must not be written to disk");
    if (process.platform !== "win32") {
        assert.equal((await fs.stat(file)).mode & 0o777, 0o600);
        assert.equal((await fs.stat(ctx.dir)).mode & 0o777, 0o700);
    }
});

test("callback with the wrong state is rejected and does not complete sign-in", async (t) => {
    const ctx = await setup({ signedIn: false });
    t.after(ctx.cleanup);
    const pending = await ctx.service.auth.startSignIn({ method: "browser" });
    const port = new URL(new URL(pending.url).searchParams.get("redirect_uri")).port;
    const res = await fetch(`http://127.0.0.1:${port}/?code=stolen&state=wrong`);
    assert.equal(res.status, 400);
    const status = await ctx.service.auth.status();
    assert.equal(status.signedIn, false);
    assert.ok(status.pending, "flow is still waiting for the real callback");
    await fakeBrowser(pending.url);
    await ctx.service.auth.getAccessToken({ waitForPendingMs: 5000 });
    assert.equal((await ctx.service.auth.status()).signedIn, true);
});

test("admin consent errors explain what to do", async (t) => {
    const ctx = await setup({ signedIn: false });
    t.after(ctx.cleanup);
    ctx.cloud.state.authorizeError = {
        error: "access_denied",
        error_description: "AADSTS90094: The grant requires admin permission.",
    };
    const pending = await ctx.service.auth.startSignIn({ method: "browser" });
    const page = await fakeBrowser(pending.url);
    assert.equal(page.status, 400);
    assert.match(page.body, /Admin approval required/);
    await assert.rejects(ctx.service.auth.getAccessToken({ waitForPendingMs: 2000 }), NotSignedInError);
    const status = await ctx.service.auth.status();
    assert.match(status.lastError, /admin approval is required/);
    assert.match(status.lastError, /Admin approval" in the README/);
});

test("a new process reuses the saved sign-in via refresh token", async (t) => {
    const ctx = await setup();
    t.after(ctx.cleanup);
    const config = await loadConfig({ env: ctx.env });
    const second = new AuthManager({ config, store: new TokenStore(config.dataDir), openUrl: async () => true });
    const token = await second.getAccessToken();
    assert.equal(token, "at-2");
    const saved = JSON.parse(await fs.readFile(path.join(ctx.dir, "auth.json"), "utf8"));
    assert.equal(saved.refreshToken, "rt-2", "rotated refresh token is saved");
});

test("concurrent refreshes share one token request", async (t) => {
    const ctx = await setup();
    t.after(ctx.cleanup);
    ctx.service.auth.invalidateAccessToken();
    const before = ctx.cloud.state.requests.length;
    const tokens = await Promise.all([1, 2, 3].map(() => ctx.service.auth.getAccessToken()));
    assert.deepEqual(new Set(tokens).size, 1);
    assert.equal(ctx.cloud.state.requests.length - before, 1);
});

test("an expired refresh token signs the user out with a clear error", async (t) => {
    const ctx = await setup();
    t.after(ctx.cleanup);
    ctx.cloud.state.refreshTokens.set("rt-1", false);
    ctx.service.auth.invalidateAccessToken();
    await assert.rejects(ctx.service.auth.getAccessToken(), (err) => err instanceof NotSignedInError && /sign in again/.test(err.message));
    assert.equal((await ctx.service.auth.status()).signedIn, false);
    await assert.rejects(fs.stat(path.join(ctx.dir, "auth.json")), { code: "ENOENT" });
});

function slowFetch(ms) {
    return async (url, init) => {
        await new Promise((r) => setTimeout(r, ms));
        return fetch(url, init);
    };
}

test("signing out during a refresh stays signed out", async (t) => {
    const ctx = await setup();
    t.after(ctx.cleanup);
    const config = await loadConfig({ env: ctx.env });
    const slow = new AuthManager({ config, store: new TokenStore(config.dataDir), fetchImpl: slowFetch(200) });
    const refreshing = slow.getAccessToken();
    await new Promise((r) => setTimeout(r, 50));
    await slow.signOut();
    await assert.rejects(refreshing, NotSignedInError);
    assert.equal((await slow.status()).signedIn, false);
    await new Promise((r) => setTimeout(r, 50));
    await assert.rejects(fs.stat(path.join(ctx.dir, "auth.json")), { code: "ENOENT" });
});

test("a refresh for the old account cannot overwrite a new sign-in", async (t) => {
    const ctx = await setup();
    t.after(ctx.cleanup);
    const config = await loadConfig({ env: ctx.env });
    let delay = 300;
    const fetchImpl = async (url, init) => {
        const ms = delay;
        await new Promise((r) => setTimeout(r, ms));
        return fetch(url, init);
    };
    const auth = new AuthManager({ config, store: new TokenStore(config.dataDir), fetchImpl, openUrl: async () => true });
    const refreshing = auth.getAccessToken();
    // Let the refresh request start (and pick up the long delay) before signing in again.
    await new Promise((r) => setTimeout(r, 20));
    delay = 0;
    const pending = await auth.startSignIn({ method: "browser", force: true });
    await fakeBrowser(pending.url);
    const fresh = await auth.getAccessToken({ waitForPendingMs: 5000 });
    const stale = await refreshing;
    assert.equal(stale, fresh, "the stale refresh returns the new token instead of committing its own");
    assert.equal(auth.access.token, fresh);
    const saved = JSON.parse(await fs.readFile(path.join(ctx.dir, "auth.json"), "utf8"));
    assert.equal(saved.refreshToken, `rt-${fresh.slice(3)}`);
});

test("a cancelled device code flow does not sign in", async (t) => {
    const ctx = await setup({ signedIn: false, authOptions: { pollIntervalMs: 10 } });
    t.after(ctx.cleanup);
    await ctx.service.auth.startSignIn({ method: "device_code" });
    await ctx.service.auth.signOut();
    await new Promise((r) => setTimeout(r, 100));
    assert.equal((await ctx.service.auth.status()).signedIn, false);
});

test("device code sign-in polls until the user finishes", async (t) => {
    const ctx = await setup({ signedIn: false, authOptions: { pollIntervalMs: 10 } });
    t.after(ctx.cleanup);
    const pending = await ctx.service.auth.startSignIn({ method: "device_code" });
    assert.equal(pending.userCode, "ABCD-1234");
    assert.equal(pending.verificationUri, "https://microsoft.com/devicelogin");
    assert.deepEqual(ctx.opened, ["https://microsoft.com/devicelogin"]);
    const token = await ctx.service.auth.getAccessToken({ waitForPendingMs: 5000 });
    assert.equal(token, "at-1");
    assert.equal(ctx.cloud.state.devicePolls, 2);
});

test("not signed in errors point to the sign-in tool", async (t) => {
    const ctx = await setup({ signedIn: false });
    t.after(ctx.cleanup);
    await assert.rejects(ctx.service.auth.getAccessToken(), /teams_chat_sign_in/);
});

test("sign-out deletes the saved token and cancels a pending flow", async (t) => {
    const ctx = await setup();
    t.after(ctx.cleanup);
    await ctx.service.auth.signOut();
    await assert.rejects(fs.stat(path.join(ctx.dir, "auth.json")), { code: "ENOENT" });
    await ctx.service.auth.startSignIn({ method: "browser" });
    await ctx.service.auth.signOut();
    assert.equal((await ctx.service.auth.status()).pending, null);
});

test("a saved sign-in for a different client ID is ignored", async (t) => {
    const ctx = await setup();
    t.after(ctx.cleanup);
    const env = { ...ctx.env, TEAMS_CHAT_CLIENT_ID: "00000000-0000-0000-0000-000000000001" };
    const config = await loadConfig({ env });
    const other = new AuthManager({ config, store: new TokenStore(config.dataDir) });
    assert.equal((await other.status()).signedIn, false);
});

test("config validation rejects bad values", async () => {
    await assert.rejects(loadConfig({ env: { TEAMS_CHAT_DATA_DIR: "/nonexistent", TEAMS_CHAT_TENANT_ID: "bad/tenant" } }), /Invalid tenantId/);
    await assert.rejects(loadConfig({ env: { TEAMS_CHAT_DATA_DIR: "/nonexistent", TEAMS_CHAT_CLIENT_ID: "nope" } }), /Invalid clientId/);
    const defaults = await loadConfig({ env: { TEAMS_CHAT_DATA_DIR: "/nonexistent" } });
    assert.equal(defaults.tenantId, "organizations");
    assert.equal(defaults.confirmBeforeSend, true);
    assert.equal(defaults.graphBaseUrl, "https://graph.microsoft.com/v1.0");
});
