import http from "node:http";
import { createHash } from "node:crypto";

function jwt(claims) {
    const enc = (o) => Buffer.from(JSON.stringify(o)).toString("base64url");
    return `${enc({ alg: "none", typ: "JWT" })}.${enc(claims)}.sig`;
}

export const ME = { id: "me-id", displayName: "Test User", mail: "test@contoso.com", userPrincipalName: "test@contoso.com" };

export const USERS = [
    ME,
    { id: "alice-id", displayName: "Alice Smith", mail: "alice@contoso.com", userPrincipalName: "alice@contoso.com", jobTitle: "PM" },
    { id: "alex-id", displayName: "Alex Smith", mail: "alex@contoso.com", userPrincipalName: "alex@contoso.com", jobTitle: "Engineer" },
    { id: "bob-id", displayName: "Bob Lee", mail: "bob@contoso.com", userPrincipalName: "bob@contoso.com" },
    { id: "carol-id", displayName: "Carol Diaz", mail: "carol@contoso.com", userPrincipalName: "carol@contoso.com" },
    // Mail differs from UPN, so GET /users/{mail} misses and search is needed.
    { id: "jane-id", displayName: "Jane Doe", mail: "jane.doe@contoso.com", userPrincipalName: "jdoe@contoso.com" },
    { id: "minh-id", displayName: "Nguyen, Minh", mail: "minh.nguyen@contoso.com", userPrincipalName: "minh.nguyen@contoso.com" },
];

const member = (u) => ({ userId: u.id, displayName: u.displayName, email: u.mail });

/**
 * A fake Microsoft identity platform + Microsoft Graph. Records every request
 * so tests can assert on what the extension actually sent.
 */
export async function startMockCloud() {
    const state = {
        requests: [],
        messages: [],
        createdChats: [],
        validAccessTokens: new Set(),
        refreshTokens: new Map(), // token -> still valid
        authCodes: new Map(), // code -> { challenge, redirectUri }
        devicePolls: 0,
        tokenCounter: 0,
        failNext: {}, // "METHOD /path" -> [status, ...]
        chats: [
            { id: "19:alice-chat@unq.gbl.spaces", chatType: "oneOnOne", topic: null, members: [member(ME), member(USERS[1])], lastUpdatedDateTime: "2026-09-26T10:00:00Z" },
            { id: "19:team-chat@thread.v2", chatType: "group", topic: null, members: [member(ME), member(USERS[3]), member(USERS[4])], lastUpdatedDateTime: "2026-09-25T10:00:00Z" },
            { id: "19:launch@thread.v2", chatType: "group", topic: "Launch plan", members: [member(ME), member(USERS[1]), member(USERS[3]), member(USERS[4])], lastUpdatedDateTime: "2026-09-24T10:00:00Z" },
        ],
        authorizeError: null, // { error, error_description } to return from /authorize
        rejectOrderBy: 0, // status to return when $orderby is used
    };

    const issue = () => {
        const n = ++state.tokenCounter;
        const access = `at-${n}`;
        const refresh = `rt-${n}`;
        state.validAccessTokens.add(access);
        state.refreshTokens.set(refresh, true);
        return {
            token_type: "Bearer",
            access_token: access,
            refresh_token: refresh,
            expires_in: 3600,
            scope: "ChatMessage.Send Chat.ReadBasic Chat.Create User.ReadBasic.All",
            id_token: jwt({ oid: ME.id, name: ME.displayName, preferred_username: ME.userPrincipalName, tid: "tenant-id" }),
        };
    };

    const readBody = async (req) => {
        const chunks = [];
        for await (const c of req) chunks.push(c);
        return Buffer.concat(chunks).toString("utf8");
    };

    const json = (res, status, body) => {
        res.writeHead(status, { "Content-Type": "application/json" });
        res.end(JSON.stringify(body));
    };

    const server = http.createServer(async (req, res) => {
        const url = new URL(req.url, "http://127.0.0.1");
        const raw = await readBody(req);
        state.requests.push({ method: req.method, path: url.pathname, search: url.search, headers: req.headers, body: raw });

        const failures = state.failNext[`${req.method} ${decodeURIComponent(url.pathname)}`];
        if (failures?.length) {
            const status = failures.shift();
            res.writeHead(status, status === 429 ? { "Retry-After": "0" } : {});
            res.end(JSON.stringify({ error: { code: "Throttled", message: "Slow down" } }));
            return;
        }

        const idp = url.pathname.match(/^\/([^/]+)\/oauth2\/v2\.0\/(authorize|token|devicecode)$/);
        if (idp) {
            const endpoint = idp[2];
            if (endpoint === "authorize") {
                const code = `code-${state.authCodes.size + 1}`;
                state.authCodes.set(code, {
                    challenge: url.searchParams.get("code_challenge"),
                    redirectUri: url.searchParams.get("redirect_uri"),
                });
                const back = new URL(url.searchParams.get("redirect_uri"));
                if (state.authorizeError) {
                    for (const [k, v] of Object.entries(state.authorizeError)) back.searchParams.set(k, v);
                    back.searchParams.set("state", url.searchParams.get("state"));
                    res.writeHead(302, { Location: back.toString() });
                    res.end();
                    return;
                }
                back.searchParams.set("code", code);
                back.searchParams.set("state", url.searchParams.get("state"));
                res.writeHead(302, { Location: back.toString() });
                res.end();
                return;
            }
            const form = new URLSearchParams(raw);
            if (endpoint === "devicecode") {
                return json(res, 200, {
                    device_code: "device-1",
                    user_code: "ABCD-1234",
                    verification_uri: "https://microsoft.com/devicelogin",
                    expires_in: 60,
                    interval: 1,
                });
            }
            const grant = form.get("grant_type");
            if (grant === "urn:ietf:params:oauth:grant-type:device_code") {
                state.devicePolls++;
                if (state.devicePolls < 2) return json(res, 400, { error: "authorization_pending" });
                return json(res, 200, issue());
            }
            if (grant === "authorization_code") {
                const entry = state.authCodes.get(form.get("code"));
                const challenge = createHash("sha256").update(form.get("code_verifier") || "").digest("base64url");
                if (!entry || entry.challenge !== challenge || entry.redirectUri !== form.get("redirect_uri")) {
                    return json(res, 400, { error: "invalid_grant", error_description: "PKCE or redirect mismatch" });
                }
                state.authCodes.delete(form.get("code"));
                return json(res, 200, issue());
            }
            if (grant === "refresh_token") {
                const rt = form.get("refresh_token");
                if (!state.refreshTokens.get(rt)) return json(res, 400, { error: "invalid_grant", error_description: "AADSTS70008: expired" });
                state.refreshTokens.set(rt, false);
                return json(res, 200, issue());
            }
            return json(res, 400, { error: "unsupported_grant_type" });
        }

        if (!url.pathname.startsWith("/v1.0/")) return json(res, 404, { error: { code: "NotFound", message: "no route" } });
        const bearer = String(req.headers.authorization || "").replace(/^Bearer /, "");
        if (!state.validAccessTokens.has(bearer)) {
            return json(res, 401, { error: { code: "InvalidAuthenticationToken", message: "Access token is invalid" } });
        }
        const p = decodeURIComponent(url.pathname.slice("/v1.0".length));

        if (req.method === "GET" && p === "/me/chats") {
            if (state.rejectOrderBy && url.searchParams.has("$orderby")) {
                return json(res, state.rejectOrderBy, { error: { code: "Forbidden", message: "orderby not allowed" } });
            }
            return json(res, 200, { value: state.chats });
        }
        let m = p.match(/^\/chats\/([^/]+)$/);
        if (req.method === "GET" && m) {
            const chat = state.chats.find((c) => c.id === m[1]);
            return chat ? json(res, 200, chat) : json(res, 404, { error: { code: "NotFound", message: "Chat not found" } });
        }
        m = p.match(/^\/users\/([^/]+)$/);
        if (req.method === "GET" && m) {
            const u = USERS.find((x) => x.id === m[1] || x.userPrincipalName.toLowerCase() === m[1].toLowerCase());
            return u ? json(res, 200, u) : json(res, 404, { error: { code: "Request_ResourceNotFound", message: "not found" } });
        }
        if (req.method === "GET" && p === "/users") {
            if (req.headers.consistencylevel !== "eventual" || url.searchParams.get("$count") !== "true") {
                return json(res, 400, { error: { code: "BadRequest", message: "ConsistencyLevel required" } });
            }
            const term = (url.searchParams.get("$search") || "").match(/displayName:([^"]*)"/)?.[1]?.trim().toLowerCase() || "";
            // Like Graph, search is tokenized prefix matching: "jane@x" matches "jane.doe@x".
            const local = term.split("@")[0];
            const found = USERS.filter((u) => u.displayName.toLowerCase().includes(term) || u.mail.toLowerCase().startsWith(local));
            return json(res, 200, { value: found });
        }
        if (req.method === "POST" && p === "/chats") {
            const body = JSON.parse(raw);
            const ids = body.members.map((mm) => mm["user@odata.bind"].match(/users\('([^']+)'\)/)[1]);
            if (body.chatType === "oneOnOne") {
                const existing = state.chats.find(
                    (c) => c.chatType === "oneOnOne" && ids.every((id) => c.members.some((x) => x.userId === id)),
                );
                if (existing) return json(res, 201, existing);
            }
            const chat = {
                id: `19:new-${state.createdChats.length + 1}@thread.v2`,
                chatType: body.chatType,
                topic: body.topic ?? null,
                members: ids.map((id) => member(USERS.find((u) => u.id === id))),
            };
            state.createdChats.push({ ...chat, requestBody: body });
            state.chats.push(chat);
            return json(res, 201, chat);
        }
        m = p.match(/^\/chats\/([^/]+)\/messages$/);
        if (req.method === "POST" && m) {
            const body = JSON.parse(raw);
            const msg = {
                id: `msg-${state.messages.length + 1}`,
                chatId: m[1],
                body: body.body,
                webUrl: `https://teams.microsoft.com/l/message/${encodeURIComponent(m[1])}`,
            };
            state.messages.push(msg);
            return json(res, 201, msg);
        }
        return json(res, 404, { error: { code: "NotFound", message: `No mock for ${req.method} ${p}` } });
    });

    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    const base = `http://127.0.0.1:${server.address().port}`;
    return {
        state,
        authorityHost: base,
        graphBaseUrl: `${base}/v1.0`,
        close: () =>
            new Promise((resolve) => {
                server.close(() => resolve());
                server.closeAllConnections();
            }),
    };
}

/** Simulates the user's browser: follows the authorize redirect back to the loopback server. */
export async function fakeBrowser(url) {
    const first = await fetch(url, { redirect: "manual" });
    const location = first.headers.get("location");
    if (!location) throw new Error("authorize did not redirect");
    const res = await fetch(location.replace("://localhost:", "://127.0.0.1:"));
    return { status: res.status, body: await res.text() };
}
