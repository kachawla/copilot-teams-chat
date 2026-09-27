import { NotSignedInError } from "./auth.mjs";

export class GraphError extends Error {
    constructor(status, code, message) {
        super(message);
        this.name = "GraphError";
        this.status = status;
        this.code = code;
    }
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function explain(status, code, message) {
    if (status === 403) {
        return `${message} (HTTP 403 ${code}). The signed-in account may lack permission, or your tenant may require admin consent for the Teams chat scopes.`;
    }
    if (status === 404) return `${message} (HTTP 404 ${code}).`;
    return `${message} (HTTP ${status} ${code}).`;
}

function memberSummary(member) {
    return {
        userId: member.userId ?? null,
        displayName: member.displayName ?? null,
        email: member.email ?? null,
    };
}

export function summarizeChat(chat, myUserId) {
    const members = (chat.members ?? []).map(memberSummary);
    const others = members.filter((m) => !myUserId || m.userId !== myUserId);
    const names = others.map((m) => m.displayName || m.email).filter(Boolean);
    let label = chat.topic || "";
    if (!label) {
        if (chat.chatType === "oneOnOne") label = names[0] || "1:1 chat";
        else label = names.length ? names.slice(0, 6).join(", ") + (names.length > 6 ? `, +${names.length - 6} more` : "") : "Untitled chat";
    }
    return {
        id: chat.id,
        chatType: chat.chatType,
        topic: chat.topic ?? null,
        label,
        members,
        lastActivity: chat.lastMessagePreview?.createdDateTime || chat.lastUpdatedDateTime || null,
        webUrl: chat.webUrl ?? null,
    };
}

export class GraphClient {
    constructor({ auth, baseUrl, fetchImpl = globalThis.fetch, maxRetries = 3, maxRetryDelayMs = 10_000 }) {
        this.auth = auth;
        this.baseUrl = baseUrl.replace(/\/+$/, "");
        this.fetch = fetchImpl;
        this.maxRetries = maxRetries;
        this.maxRetryDelayMs = maxRetryDelayMs;
    }

    async request(method, pathOrUrl, { body, headers = {}, waitForSignInMs = 0 } = {}) {
        const url = /^https?:\/\//.test(pathOrUrl) ? pathOrUrl : `${this.baseUrl}${pathOrUrl}`;
        if (!url.startsWith(`${this.baseUrl}/`)) {
            throw new GraphError(0, "invalid_url", `Refusing to call a non-Graph URL: ${url}`);
        }
        let refreshedAfter401 = false;
        for (let attempt = 0; ; attempt++) {
            const token = await this.auth.getAccessToken({ waitForPendingMs: waitForSignInMs });
            const response = await this.fetch(url, {
                method,
                headers: {
                    Authorization: `Bearer ${token}`,
                    Accept: "application/json",
                    ...(body !== undefined ? { "Content-Type": "application/json" } : {}),
                    ...headers,
                },
                body: body !== undefined ? JSON.stringify(body) : undefined,
            });

            if (response.status === 401 && !refreshedAfter401) {
                refreshedAfter401 = true;
                this.auth.invalidateAccessToken();
                continue;
            }
            // A 429 means Graph did not process the request, so it is always safe to retry.
            // 503/504 may come after a POST took effect, so only reads retry on those.
            const retryable = response.status === 429 || (method === "GET" && (response.status === 503 || response.status === 504));
            if (retryable && attempt < this.maxRetries) {
                const retryAfter = Number(response.headers.get("retry-after"));
                const delay = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 1000 * 2 ** attempt;
                await sleep(Math.min(delay, this.maxRetryDelayMs));
                continue;
            }

            const text = await response.text();
            let json = null;
            if (text) {
                try {
                    json = JSON.parse(text);
                } catch {
                    json = null;
                }
            }
            if (!response.ok) {
                const code = json?.error?.code || "error";
                const message = json?.error?.message || text.slice(0, 300) || response.statusText || "Request failed";
                if (method !== "GET" && (response.status === 503 || response.status === 504)) {
                    throw new GraphError(
                        response.status,
                        code,
                        `Microsoft Graph did not respond in time (HTTP ${response.status}). The request may still have gone through; check Teams before trying again.`,
                    );
                }
                if (response.status === 401) {
                    throw new NotSignedInError(`Microsoft Graph rejected the sign-in (${message}). Call teams_chat_sign_in to sign in again.`);
                }
                throw new GraphError(response.status, code, explain(response.status, code, message));
            }
            return json;
        }
    }

    async listChats({ maxPages = 1, pageSize = 50, myUserId } = {}) {
        const query = `$expand=members&$top=${Math.min(Math.max(pageSize, 1), 50)}&$orderby=lastMessagePreview/createdDateTime desc`;
        let next = `/me/chats?${query}`;
        let page;
        try {
            page = await this.request("GET", next);
        } catch (err) {
            // Sorting by lastMessagePreview can be refused (400) or need more than Chat.ReadBasic (403).
            if (!(err instanceof GraphError) || (err.status !== 400 && err.status !== 403)) throw err;
            next = `/me/chats?$expand=members&$top=${Math.min(Math.max(pageSize, 1), 50)}`;
            page = await this.request("GET", next);
        }
        const chats = [];
        for (let i = 0; ; i++) {
            chats.push(...(page?.value ?? []).map((chat) => summarizeChat(chat, myUserId)));
            next = page?.["@odata.nextLink"];
            if (!next || i + 1 >= maxPages) break;
            page = await this.request("GET", next);
        }
        return chats;
    }

    async getChat(chatId, myUserId) {
        const chat = await this.request("GET", `/chats/${encodeURIComponent(chatId)}?$expand=members`);
        return summarizeChat(chat, myUserId);
    }

    async getUser(idOrUpn) {
        try {
            return await this.request(
                "GET",
                `/users/${encodeURIComponent(idOrUpn)}?$select=id,displayName,mail,userPrincipalName,jobTitle`,
            );
        } catch (err) {
            if (err instanceof GraphError && err.status === 404) return null;
            throw err;
        }
    }

    async searchUsers(query, { top = 10 } = {}) {
        const term = String(query).replace(/["\\]/g, " ").trim();
        if (!term) return [];
        const search = `"displayName:${term}" OR "mail:${term}"`;
        const path = `/users?$search=${encodeURIComponent(search)}&$select=id,displayName,mail,userPrincipalName,jobTitle&$top=${top}&$count=true`;
        const res = await this.request("GET", path, { headers: { ConsistencyLevel: "eventual" } });
        return res?.value ?? [];
    }

    #member(userId) {
        return {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `${this.baseUrl}/users('${userId}')`,
        };
    }

    /** Returns the existing 1:1 chat with this user, or creates it. */
    async createOneOnOneChat(myUserId, userId) {
        return this.request("POST", "/chats", {
            body: { chatType: "oneOnOne", members: [this.#member(myUserId), this.#member(userId)] },
        });
    }

    async createGroupChat(myUserId, userIds, topic) {
        return this.request("POST", "/chats", {
            body: {
                chatType: "group",
                ...(topic ? { topic } : {}),
                members: [this.#member(myUserId), ...userIds.map((id) => this.#member(id))],
            },
        });
    }

    async sendChatMessage(chatId, html) {
        return this.request("POST", `/chats/${encodeURIComponent(chatId)}/messages`, {
            body: { body: { contentType: "html", content: html } },
        });
    }
}
