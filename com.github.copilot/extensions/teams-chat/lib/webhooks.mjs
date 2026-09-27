import fs from "node:fs/promises";
import path from "node:path";
import { randomBytes } from "node:crypto";

// Teams Workflows ("When a Teams webhook request is received") and legacy
// incoming webhook hosts. Anything else is refused so a prompt-injected agent
// cannot register an arbitrary URL and use it to send data elsewhere.
const DEFAULT_HOST_SUFFIXES = [
    ".logic.azure.com",
    ".powerplatform.com",
    ".logic.azure.us",
    ".powerplatform.us",
    ".webhook.office.com",
];
const NAME = /^[A-Za-z0-9][A-Za-z0-9 ._-]{0,63}$/;
export const MAX_WEBHOOK_PAYLOAD_BYTES = 28_000;

export class WebhookError extends Error {
    constructor(message, code = "webhook_error") {
        super(message);
        this.name = "WebhookError";
        this.code = code;
    }
}

export function validateWebhookName(name) {
    const value = String(name ?? "").trim();
    if (!NAME.test(value)) {
        throw new WebhookError(`Invalid webhook name "${value}". Use up to 64 letters, numbers, spaces, dots, dashes, or underscores.`, "invalid_name");
    }
    return value;
}

export function validateWebhookUrl(raw, extraHostSuffixes = []) {
    let url;
    try {
        url = new URL(String(raw).trim());
    } catch {
        throw new WebhookError("That is not a valid URL.", "invalid_url");
    }
    if (url.protocol !== "https:") throw new WebhookError("Webhook URLs must use https.", "invalid_url");
    if (url.username || url.password) throw new WebhookError("Webhook URLs must not contain credentials.", "invalid_url");
    const host = url.hostname.toLowerCase();
    const suffixes = [...DEFAULT_HOST_SUFFIXES, ...extraHostSuffixes.map((s) => `.${String(s).toLowerCase().replace(/^\.+/, "")}`)];
    if (!suffixes.some((suffix) => host.endsWith(suffix) && host.length > suffix.length)) {
        throw new WebhookError(
            `"${host}" is not a Teams Workflows webhook host. Copy the URL from the Teams workflow "When a Teams webhook request is received" (it ends in ${DEFAULT_HOST_SUFFIXES.slice(0, 2).join(" or ")}).`,
            "untrusted_host",
        );
    }
    // Keep the exact string: the query carries a signature that must not be re-encoded.
    return String(raw).trim();
}

/** Named webhook destinations. The URLs are secrets, so the file is owner-only and URLs are never shown back. */
export class WebhookStore {
    constructor(dir, { extraHostSuffixes = [] } = {}) {
        this.dir = dir;
        this.file = path.join(dir, "webhooks.json");
        this.extraHostSuffixes = extraHostSuffixes;
    }

    async #readAll() {
        try {
            const data = JSON.parse(await fs.readFile(this.file, "utf8"));
            return data && typeof data === "object" && data.webhooks && typeof data.webhooks === "object" ? data.webhooks : {};
        } catch (err) {
            if (err.code === "ENOENT" || err instanceof SyntaxError) return {};
            throw err;
        }
    }

    async #writeAll(webhooks) {
        await fs.mkdir(this.dir, { recursive: true, mode: 0o700 });
        await fs.chmod(this.dir, 0o700).catch(() => {});
        const tmp = `${this.file}.${randomBytes(6).toString("hex")}.tmp`;
        await fs.writeFile(tmp, JSON.stringify({ webhooks }, null, 2), { mode: 0o600 });
        await fs.rename(tmp, this.file);
    }

    async list() {
        const all = await this.#readAll();
        return Object.entries(all)
            .map(([name, w]) => ({ name, description: w.description ?? null, addedAt: w.addedAt ?? null }))
            .sort((a, b) => a.name.localeCompare(b.name));
    }

    async get(name) {
        const all = await this.#readAll();
        const key = Object.keys(all).find((k) => k.toLowerCase() === String(name).trim().toLowerCase());
        return key ? { name: key, ...all[key] } : null;
    }

    async add(name, url, description) {
        const cleanName = validateWebhookName(name);
        const cleanUrl = validateWebhookUrl(url, this.extraHostSuffixes);
        const all = await this.#readAll();
        const existing = Object.keys(all).find((k) => k.toLowerCase() === cleanName.toLowerCase());
        if (existing) delete all[existing];
        all[cleanName] = {
            url: cleanUrl,
            description: description ? String(description).slice(0, 200) : null,
            addedAt: new Date().toISOString(),
        };
        await this.#writeAll(all);
        return { name: cleanName, replaced: Boolean(existing) };
    }

    async remove(name) {
        const all = await this.#readAll();
        const key = Object.keys(all).find((k) => k.toLowerCase() === String(name).trim().toLowerCase());
        if (!key) return false;
        delete all[key];
        await this.#writeAll(all);
        return true;
    }
}

/**
 * Builds an Adaptive Card for the message. TextBlocks render a Markdown
 * subset (bold, italic, links, lists); headings and code blocks become their
 * own styled blocks. Adaptive Card text is never interpreted as HTML.
 */
export function buildAdaptiveCard(message, format = "markdown") {
    const text = String(message).replace(/\r\n?/g, "\n");
    const body = [];
    const pushText = (lines) => {
        const value = lines.join("\n").trim();
        if (value) body.push({ type: "TextBlock", text: value, wrap: true });
    };

    if (format === "text") {
        pushText([text]);
    } else {
        const lines = text.split("\n");
        let buffer = [];
        for (let i = 0; i < lines.length; i++) {
            const fence = lines[i].match(/^\s*(```|~~~)/);
            if (fence) {
                pushText(buffer);
                buffer = [];
                const code = [];
                i++;
                while (i < lines.length && !lines[i].trim().startsWith(fence[1])) code.push(lines[i++]);
                body.push({ type: "TextBlock", text: code.join("\n") || " ", wrap: true, fontType: "Monospace" });
                continue;
            }
            const heading = lines[i].match(/^\s*#{1,6}\s+(.*?)\s*#*\s*$/);
            if (heading) {
                pushText(buffer);
                buffer = [];
                body.push({ type: "TextBlock", text: heading[1], wrap: true, weight: "Bolder", size: "Medium" });
                continue;
            }
            buffer.push(lines[i]);
        }
        pushText(buffer);
    }
    if (!body.length) throw new WebhookError("The message is empty.", "empty");

    return {
        type: "message",
        attachments: [
            {
                contentType: "application/vnd.microsoft.card.adaptive",
                contentUrl: null,
                content: {
                    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                    type: "AdaptiveCard",
                    version: "1.4",
                    msteams: { width: "Full" },
                    body,
                },
            },
        ],
    };
}

export function renderWebhookPayload(message, format) {
    const payload = buildAdaptiveCard(message, format);
    const bytes = new TextEncoder().encode(JSON.stringify(payload)).length;
    if (bytes > MAX_WEBHOOK_PAYLOAD_BYTES) {
        throw new WebhookError(`The message is too long for a Teams webhook (${bytes} bytes, limit ${MAX_WEBHOOK_PAYLOAD_BYTES}). Shorten it or split it.`, "too_long");
    }
    return payload;
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export async function postWebhook(url, payload, { fetchImpl = globalThis.fetch, maxRetries = 2 } = {}) {
    for (let attempt = 0; ; attempt++) {
        let response;
        try {
            response = await fetchImpl(url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
                redirect: "error",
            });
        } catch (err) {
            throw new WebhookError(`Could not reach the Teams webhook: ${err.message}`, "network");
        }
        if (response.ok) return { status: response.status };
        // 429 means the request was not accepted, so a retry cannot double-post.
        if (response.status === 429 && attempt < maxRetries) {
            const retryAfter = Number(response.headers.get("retry-after"));
            await sleep(Math.min(Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 1000 * 2 ** attempt, 10_000));
            continue;
        }
        const detail = (await response.text().catch(() => "")).slice(0, 300);
        if (response.status === 401 || response.status === 403 || response.status === 404) {
            throw new WebhookError(
                `The Teams webhook rejected the message (HTTP ${response.status}). The workflow may be turned off, deleted, or its URL changed. Recreate it in Teams and add the new URL.`,
                "rejected",
            );
        }
        if (response.status >= 500) {
            throw new WebhookError(
                `The Teams webhook failed (HTTP ${response.status}). The message may still have been posted; check Teams before trying again.`,
                "server_error",
            );
        }
        throw new WebhookError(`The Teams webhook rejected the message (HTTP ${response.status}${detail ? `: ${detail}` : ""}).`, "rejected");
    }
}
