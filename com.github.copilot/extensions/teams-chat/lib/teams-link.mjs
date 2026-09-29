// Builds Microsoft Teams "start a chat" deep links and turns Markdown into
// plain text for the Teams compose box, which doesn't render Markdown from a
// link. Browser-compatible: the compose canvas imports this file for its preview.
// Deep link format: https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/deep-link-teams

export const WEB_BASE = "https://teams.microsoft.com/l/chat/0/0";
export const APP_BASE = "msteams:/l/chat/0/0";

// Long links can be cut off by the OS or Teams. Kept well under common limits.
export const MAX_LINK_LENGTH = 8000;

export class LinkError extends Error {
    constructor(message, code = "link_error") {
        super(message);
        this.name = "LinkError";
        this.code = code;
    }
}

function stripInline(text) {
    return text
        .replace(/!\[([^\]\n]*)\]\(([^)\s]+)\)/g, "$1 ($2)")
        .replace(/\[([^\]\n]+)\]\(([^)\s]+)\)/g, (_, label, url) => (label.trim() === url.trim() ? url : `${label} (${url})`))
        .replace(/<(https?:\/\/[^>\s]+)>/g, "$1")
        .replace(/`([^`\n]+)`/g, "$1")
        .replace(/\*\*([^*\n]+)\*\*/g, "$1")
        .replace(/__([^_\n]+)__/g, "$1")
        .replace(/~~([^~\n]+)~~/g, "$1")
        .replace(/(^|[^*\w])\*([^*\s][^*\n]*?)\*(?!\w)/g, "$1$2")
        .replace(/(^|[^_\w])_([^_\s][^_\n]*?)_(?!\w)/g, "$1$2");
}

/** Converts Markdown to readable plain text: keeps structure, drops syntax. */
export function markdownToPlainText(markdown) {
    const lines = String(markdown).replace(/\r\n?/g, "\n").split("\n");
    const out = [];
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const fence = line.match(/^\s*(```|~~~)/);
        if (fence) {
            i++;
            while (i < lines.length && !lines[i].trim().startsWith(fence[1])) out.push(lines[i++]);
            continue;
        }
        if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(line)) {
            out.push("");
            continue;
        }
        const heading = line.match(/^\s*#{1,6}\s+(.*?)\s*#*\s*$/);
        if (heading) {
            out.push(stripInline(heading[1]));
            continue;
        }
        const bullet = line.match(/^(\s*)[-*+]\s+(.*)$/);
        if (bullet) {
            out.push(`${bullet[1]}• ${stripInline(bullet[2])}`);
            continue;
        }
        const quote = line.match(/^\s*>\s?(.*)$/);
        if (quote) {
            out.push(stripInline(quote[1]));
            continue;
        }
        out.push(stripInline(line));
    }
    return out.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

export function toTeamsText(message, format = "markdown") {
    const text = format === "text" ? String(message).replace(/\r\n?/g, "\n").trim() : markdownToPlainText(message);
    if (!text) throw new LinkError("The message is empty.", "empty");
    return text;
}

/**
 * Builds the deep link. One person opens your existing 1:1 chat with them;
 * several people open the chat with exactly those members, or a new draft
 * group chat named `topic`. Nothing is sent until the user presses Send.
 */
export function buildChatLink({ emails, topic, text, tenantId, app = false }) {
    if (!emails?.length) throw new LinkError("Say who the message is for.", "no_recipients");
    const params = [`users=${emails.map((e) => encodeURIComponent(e)).join(",")}`];
    if (tenantId) params.push(`tenantId=${encodeURIComponent(tenantId)}`);
    if (topic && emails.length > 1) params.push(`topicName=${encodeURIComponent(String(topic).trim().slice(0, 250))}`);
    if (text) params.push(`message=${encodeURIComponent(text)}`);
    const link = `${app ? APP_BASE : WEB_BASE}?${params.join("&")}`;
    if (link.length > MAX_LINK_LENGTH) {
        throw new LinkError(`The message is too long to open in Teams (${link.length} characters in the link, limit ${MAX_LINK_LENGTH}). Shorten it or split it.`, "too_long");
    }
    return link;
}
