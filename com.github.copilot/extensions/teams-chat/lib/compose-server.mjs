import fs from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { randomBytes, timingSafeEqual } from "node:crypto";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const STATIC_FILES = {
    "/": { file: path.join(here, "..", "ui", "compose.html"), type: "text/html; charset=utf-8" },
    "/app.js": { file: path.join(here, "..", "ui", "app.js"), type: "text/javascript; charset=utf-8" },
    "/style.css": { file: path.join(here, "..", "ui", "style.css"), type: "text/css; charset=utf-8" },
    "/markdown.mjs": { file: path.join(here, "markdown.mjs"), type: "text/javascript; charset=utf-8" },
};
const MAX_BODY_BYTES = 256 * 1024;
const CSP = "default-src 'none'; script-src 'self'; style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'none'";

function safeEqual(a, b) {
    const x = Buffer.from(String(a));
    const y = Buffer.from(String(b));
    return x.length === y.length && timingSafeEqual(x, y);
}

async function readJson(req) {
    const chunks = [];
    let size = 0;
    for await (const chunk of req) {
        size += chunk.length;
        if (size > MAX_BODY_BYTES) throw Object.assign(new Error("Request body too large."), { status: 413 });
        chunks.push(chunk);
    }
    if (!size) return {};
    try {
        return JSON.parse(Buffer.concat(chunks).toString("utf8"));
    } catch {
        throw Object.assign(new Error("Invalid JSON body."), { status: 400 });
    }
}

function sendJson(res, status, value) {
    res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
    res.end(JSON.stringify(value));
}

/**
 * Loopback HTTP server for one compose canvas panel. API calls must carry the
 * per-panel token (header, or query for EventSource) and a loopback Host
 * header, which blocks cross-site requests and DNS rebinding.
 */
export async function startComposeServer({ getState, onDraft, onSend }) {
    const token = randomBytes(24).toString("base64url");
    const clients = new Set();
    let port = 0;

    const server = http.createServer(async (req, res) => {
        try {
            const host = String(req.headers.host || "");
            if (host !== `127.0.0.1:${port}` && host !== `localhost:${port}`) {
                res.writeHead(421).end();
                return;
            }
            const url = new URL(req.url, `http://127.0.0.1:${port}`);
            const asset = STATIC_FILES[url.pathname];
            if (req.method === "GET" && asset) {
                const body = await fs.readFile(asset.file);
                res.writeHead(200, {
                    "Content-Type": asset.type,
                    "Content-Security-Policy": CSP,
                    "Cache-Control": "no-store",
                    "X-Content-Type-Options": "nosniff",
                    "Referrer-Policy": "no-referrer",
                });
                res.end(body);
                return;
            }

            if (!url.pathname.startsWith("/api/")) {
                res.writeHead(404).end();
                return;
            }
            const presented = req.headers["x-teams-chat-token"] ?? url.searchParams.get("token") ?? "";
            if (!safeEqual(presented, token)) {
                sendJson(res, 403, { error: "Forbidden" });
                return;
            }

            if (req.method === "GET" && url.pathname === "/api/state") {
                sendJson(res, 200, await getState());
                return;
            }
            if (req.method === "GET" && url.pathname === "/api/events") {
                res.writeHead(200, {
                    "Content-Type": "text/event-stream",
                    "Cache-Control": "no-store",
                    Connection: "keep-alive",
                });
                res.write(`data: ${JSON.stringify(await getState())}\n\n`);
                clients.add(res);
                req.on("close", () => clients.delete(res));
                return;
            }
            // Mutations require the header token so a plain cross-site form post cannot trigger them.
            if (req.method === "POST" && !safeEqual(req.headers["x-teams-chat-token"] ?? "", token)) {
                sendJson(res, 403, { error: "Forbidden" });
                return;
            }
            if (req.method === "POST" && url.pathname === "/api/draft") {
                sendJson(res, 200, await onDraft(await readJson(req)));
                return;
            }
            if (req.method === "POST" && url.pathname === "/api/send") {
                sendJson(res, 200, await onSend(await readJson(req)));
                return;
            }
            sendJson(res, 404, { error: "Not found" });
        } catch (err) {
            if (!res.headersSent) sendJson(res, err.status || 500, { error: err.message });
            else res.end();
        }
    });

    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    port = server.address().port;

    return {
        url: `http://127.0.0.1:${port}/?token=${token}`,
        token,
        port,
        async broadcast() {
            if (!clients.size) return;
            const payload = `data: ${JSON.stringify(await getState())}\n\n`;
            for (const client of clients) client.write(payload);
        },
        close() {
            for (const client of clients) client.end();
            clients.clear();
            return new Promise((resolve) => {
                server.close(() => resolve());
                server.closeAllConnections?.();
            });
        },
    };
}
