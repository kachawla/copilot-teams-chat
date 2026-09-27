import http from "node:http";
import { spawn } from "node:child_process";
import { createHash, randomBytes } from "node:crypto";
import { EventEmitter } from "node:events";
import { GRAPH_SCOPES, OIDC_SCOPES } from "./config.mjs";

export class NotSignedInError extends Error {
    constructor(message = "Not signed in to Microsoft Teams. Call teams_chat_sign_in first.") {
        super(message);
        this.name = "NotSignedInError";
        this.code = "not_signed_in";
    }
}

export class AuthError extends Error {
    constructor(message, code = "auth_error") {
        super(message);
        this.name = "AuthError";
        this.code = code;
    }
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export function decodeJwtClaims(token) {
    try {
        const payload = String(token).split(".")[1];
        return JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) || {};
    } catch {
        return {};
    }
}

function describeTokenError(res) {
    const description = String(res.error_description || "").split(/\r?\n/)[0];
    return description ? `${res.error}: ${description}` : String(res.error);
}

function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

function callbackPage(title, detail) {
    return `<!doctype html><meta charset="utf-8"><title>${escapeHtml(title)}</title>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:32rem;margin:4rem auto;padding:0 1rem">
<h1 style="font-size:1.4rem">${escapeHtml(title)}</h1><p>${escapeHtml(detail)}</p></body>`;
}

export function openInBrowser(url) {
    let command;
    let args;
    if (process.platform === "darwin") {
        command = "open";
        args = [url];
    } else if (process.platform === "win32") {
        command = "rundll32";
        args = ["url.dll,FileProtocolHandler", url];
    } else {
        command = "xdg-open";
        args = [url];
    }
    return new Promise((resolve) => {
        try {
            const child = spawn(command, args, { stdio: "ignore", detached: true });
            child.once("error", () => resolve(false));
            child.once("spawn", () => {
                child.unref();
                resolve(true);
            });
        } catch {
            resolve(false);
        }
    });
}

function listen(server, port, host) {
    return new Promise((resolve, reject) => {
        server.once("error", reject);
        server.listen(port, host, () => {
            server.off("error", reject);
            resolve();
        });
    });
}

/**
 * Delegated Microsoft identity platform sign-in without third-party
 * dependencies: authorization code + PKCE on a loopback redirect, or the
 * device code flow for hosts without a local browser.
 *
 * Events: "change" (any auth state change), "signed-in" (account),
 * "sign-in-failed" (error).
 */
export class AuthManager extends EventEmitter {
    constructor({
        config,
        store,
        fetchImpl = globalThis.fetch,
        openUrl = openInBrowser,
        now = Date.now,
        browserTimeoutMs = 5 * 60_000,
        pollIntervalMs,
    }) {
        super();
        this.config = config;
        this.store = store;
        this.fetch = fetchImpl;
        this.openUrl = openUrl;
        this.now = now;
        this.browserTimeoutMs = browserTimeoutMs;
        this.pollIntervalMs = pollIntervalMs;
        this.saved = null;
        this.access = null;
        this.pending = null;
        this.refreshing = null;
        this.lastError = null;
        // Bumped whenever the signed-in identity changes (sign-out, cancelled
        // flow, new sign-in). Token responses from an older generation are dropped.
        this.generation = 0;
        // Serializes auth.json writes and deletes so disk matches memory.
        this.persisting = Promise.resolve();
        this.ready = this.#load();
    }

    get scope() {
        return [...GRAPH_SCOPES, ...OIDC_SCOPES].join(" ");
    }

    get authorityBase() {
        return `${this.config.authorityHost}/${this.config.tenantId}/oauth2/v2.0`;
    }

    async #load() {
        const saved = await this.store.read().catch(() => null);
        if (saved && saved.clientId === this.config.clientId && saved.tenantId === this.config.tenantId) {
            this.saved = saved;
        }
    }

    async status() {
        await this.ready;
        return {
            signedIn: Boolean(this.saved),
            account: this.saved?.account ?? null,
            pending: this.describePending(),
            lastError: this.lastError,
            clientId: this.config.clientId,
            tenantId: this.config.tenantId,
        };
    }

    describePending() {
        if (!this.pending) return null;
        const { method, url, userCode, verificationUri, browserOpened, startedAt } = this.pending;
        return { method, url, userCode, verificationUri, browserOpened, startedAt };
    }

    async getAccessToken({ waitForPendingMs = 0 } = {}) {
        await this.ready;
        // Hand out a token only once the sign-in that produced it is saved.
        await this.persisting;
        if (this.#accessValid()) return this.access.token;
        if (!this.saved && this.pending && waitForPendingMs > 0) {
            await Promise.race([this.pending.done.catch(() => {}), sleep(waitForPendingMs)]);
            await this.persisting;
            if (this.#accessValid()) return this.access.token;
        }
        if (!this.saved) {
            throw new NotSignedInError(
                this.pending
                    ? "Microsoft Teams sign-in is still in progress. Finish signing in, then try again."
                    : undefined,
            );
        }
        return this.refresh();
    }

    #accessValid() {
        return Boolean(this.access && this.access.expiresAt - 60_000 > this.now());
    }

    invalidateAccessToken() {
        this.access = null;
    }

    refresh() {
        if (!this.refreshing) {
            this.refreshing = this.#refresh().finally(() => {
                this.refreshing = null;
            });
        }
        return this.refreshing;
    }

    async #refresh() {
        if (!this.saved?.refreshToken) {
            throw new NotSignedInError("Your Microsoft Teams sign-in expired. Call teams_chat_sign_in to sign in again.");
        }
        const generation = this.generation;
        const res = await this.#tokenRequest({
            grant_type: "refresh_token",
            refresh_token: this.saved.refreshToken,
            scope: this.scope,
        });
        if (generation !== this.generation) {
            // Signed out or switched accounts while refreshing.
            if (this.#accessValid()) return this.access.token;
            throw new NotSignedInError();
        }
        if (res.error) {
            if (res.error === "invalid_grant" || res.error === "interaction_required") {
                await this.signOut();
                throw new NotSignedInError(
                    `Your Microsoft Teams sign-in expired (${describeTokenError(res)}). Call teams_chat_sign_in to sign in again.`,
                );
            }
            throw new AuthError(`Token refresh failed: ${describeTokenError(res)}`, res.error);
        }
        await this.#acceptTokens(res, { generation });
        return this.access.token;
    }

    async #tokenRequest(params) {
        const body = new URLSearchParams({ client_id: this.config.clientId, ...params });
        const response = await this.fetch(`${this.authorityBase}/token`, {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body,
        });
        let json;
        try {
            json = await response.json();
        } catch {
            json = { error: "invalid_response", error_description: `HTTP ${response.status}` };
        }
        if (!response.ok && !json.error) json.error = `http_${response.status}`;
        return json;
    }

    #persist(op) {
        const next = this.persisting.then(op, op);
        this.persisting = next.catch(() => {});
        return next;
    }

    async #acceptTokens(res, { generation, flow = null }) {
        if (!res.access_token) throw new AuthError("The token response did not include an access token.");
        if (generation !== this.generation || (flow && this.pending !== flow)) {
            throw new AuthError("Sign-in cancelled.", "cancelled");
        }
        const id = res.id_token ? decodeJwtClaims(res.id_token) : {};
        const at = decodeJwtClaims(res.access_token);
        const previous = flow ? {} : (this.saved?.account ?? {});
        const account = {
            id: id.oid || at.oid || previous.id || null,
            displayName: id.name || at.name || previous.displayName || null,
            username: id.preferred_username || at.upn || at.unique_name || previous.username || null,
            tenantId: id.tid || at.tid || previous.tenantId || null,
        };
        const saved = {
            refreshToken: res.refresh_token || (flow ? null : this.saved?.refreshToken) || null,
            account,
            clientId: this.config.clientId,
            tenantId: this.config.tenantId,
            savedAt: new Date(this.now()).toISOString(),
        };
        // Commit synchronously after the generation check so no other auth
        // change can interleave; the disk write is queued in the same order.
        const wasSignedIn = Boolean(this.saved);
        if (flow) {
            this.generation++;
            this.pending = null;
        }
        this.saved = saved;
        this.access = {
            token: res.access_token,
            expiresAt: this.now() + (Number(res.expires_in) || 3600) * 1000,
        };
        this.lastError = null;
        const write = saved.refreshToken ? this.#persist(() => this.store.write(saved)) : Promise.resolve();
        if (flow || !wasSignedIn) this.emit("signed-in", account);
        this.emit("change");
        await write;
    }

    /** Starts a sign-in flow and returns immediately with what the user needs to finish it. */
    async startSignIn({ method = "browser", force = false } = {}) {
        await this.ready;
        if (this.pending) {
            if (!force && this.pending.method === method) return this.describePending();
            this.#cancelPending();
        }
        this.lastError = null;
        const flow = method === "device_code" ? await this.#startDeviceCode() : await this.#startBrowser();
        // Another sign-in may have started while this one was being set up.
        if (this.pending) this.#cancelPending();
        const generation = this.generation;
        flow.done = flow.tokens.then((tokens) => this.#acceptTokens(tokens, { generation, flow }));
        this.pending = flow;
        flow.done.then(
            () => {},
            (err) => {
                if (this.pending !== flow) return;
                this.pending = null;
                this.lastError = err.message;
                this.emit("sign-in-failed", err);
                this.emit("change");
            },
        );
        this.emit("change");
        return this.describePending();
    }

    #cancelPending() {
        const flow = this.pending;
        if (!flow) return;
        this.pending = null;
        this.generation++;
        flow.cancel();
    }

    async #startBrowser() {
        const verifier = randomBytes(32).toString("base64url");
        const challenge = createHash("sha256").update(verifier).digest("base64url");
        const state = randomBytes(16).toString("base64url");
        let redirectUri = "";
        let settle;
        let settled = false;
        const tokens = new Promise((resolve, reject) => {
            settle = {
                resolve: (value) => {
                    settled = true;
                    resolve(value);
                },
                reject: (err) => {
                    settled = true;
                    reject(err);
                },
            };
        });

        const handler = (req, res) => {
            const url = new URL(req.url, "http://localhost");
            const params = url.searchParams;
            if (url.pathname !== "/" || (!params.has("code") && !params.has("error"))) {
                res.writeHead(404).end();
                return;
            }
            const send = (status, title, detail) => {
                res.writeHead(status, { "Content-Type": "text/html; charset=utf-8", Connection: "close" });
                res.end(callbackPage(title, detail));
            };
            if (params.get("state") !== state || settled) {
                send(400, "Sign-in request not recognized", "Start sign-in again from GitHub Copilot.");
                return;
            }
            if (params.has("error")) {
                const err = describeTokenError({ error: params.get("error"), error_description: params.get("error_description") });
                send(400, "Sign-in failed", err);
                settle.reject(new AuthError(`Sign-in failed: ${err}`, params.get("error")));
                return;
            }
            settled = true;
            this.#tokenRequest({
                grant_type: "authorization_code",
                code: params.get("code"),
                redirect_uri: redirectUri,
                code_verifier: verifier,
                scope: this.scope,
            }).then(
                (res2) => {
                    if (res2.error) {
                        send(400, "Sign-in failed", describeTokenError(res2));
                        settle.reject(new AuthError(`Sign-in failed: ${describeTokenError(res2)}`, res2.error));
                    } else {
                        send(200, "Signed in to Microsoft Teams", "You can close this tab and return to GitHub Copilot.");
                        settle.resolve(res2);
                    }
                },
                (err) => {
                    send(500, "Sign-in failed", err.message);
                    settle.reject(err);
                },
            );
        };

        const servers = [http.createServer(handler)];
        await listen(servers[0], 0, "127.0.0.1");
        const port = servers[0].address().port;
        // Browsers may resolve "localhost" to ::1 first, so listen there too when available.
        const v6 = http.createServer(handler);
        await listen(v6, port, "::1").then(
            () => servers.push(v6),
            () => {},
        );
        redirectUri = `http://localhost:${port}`;

        const timer = setTimeout(
            () => settle.reject(new AuthError("Sign-in timed out. Call teams_chat_sign_in to try again.", "timeout")),
            this.browserTimeoutMs,
        );
        const cleanup = () => {
            clearTimeout(timer);
            setTimeout(() => {
                for (const server of servers) {
                    server.close();
                    server.closeAllConnections?.();
                }
            }, 250).unref();
        };
        tokens.then(cleanup, cleanup);

        const authorize = new URL(`${this.authorityBase}/authorize`);
        authorize.search = new URLSearchParams({
            client_id: this.config.clientId,
            response_type: "code",
            redirect_uri: redirectUri,
            response_mode: "query",
            scope: this.scope,
            state,
            code_challenge: challenge,
            code_challenge_method: "S256",
            prompt: "select_account",
        }).toString();
        const url = authorize.toString();
        const browserOpened = await this.openUrl(url).catch(() => false);

        return {
            method: "browser",
            url,
            browserOpened,
            startedAt: new Date(this.now()).toISOString(),
            tokens,
            cancel: () => settle.reject(new AuthError("Sign-in cancelled.", "cancelled")),
        };
    }

    async #startDeviceCode() {
        const response = await this.fetch(`${this.authorityBase}/devicecode`, {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({ client_id: this.config.clientId, scope: this.scope }),
        });
        const dc = await response.json().catch(() => ({ error: `http_${response.status}` }));
        if (!response.ok || dc.error) {
            throw new AuthError(`Could not start device code sign-in: ${describeTokenError(dc)}`, dc.error);
        }

        let cancelled = false;
        const tokens = (async () => {
            let interval = this.pollIntervalMs ?? (Number(dc.interval) || 5) * 1000;
            const deadline = this.now() + (Number(dc.expires_in) || 900) * 1000;
            while (!cancelled && this.now() < deadline) {
                await sleep(interval);
                if (cancelled) break;
                const res = await this.#tokenRequest({
                    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
                    device_code: dc.device_code,
                });
                if (!res.error) return res;
                if (res.error === "authorization_pending") continue;
                if (res.error === "slow_down") {
                    interval += 5000;
                    continue;
                }
                throw new AuthError(`Sign-in failed: ${describeTokenError(res)}`, res.error);
            }
            throw cancelled
                ? new AuthError("Sign-in cancelled.", "cancelled")
                : new AuthError("The sign-in code expired. Call teams_chat_sign_in to try again.", "timeout");
        })();

        const browserOpened = await this.openUrl(dc.verification_uri).catch(() => false);
        return {
            method: "device_code",
            userCode: dc.user_code,
            verificationUri: dc.verification_uri,
            browserOpened,
            startedAt: new Date(this.now()).toISOString(),
            tokens,
            cancel: () => {
                cancelled = true;
            },
        };
    }

    async signOut() {
        await this.ready;
        this.generation++;
        this.#cancelPending();
        this.access = null;
        this.saved = null;
        const clear = this.#persist(() => this.store.clear());
        this.emit("change");
        await clear;
    }
}
