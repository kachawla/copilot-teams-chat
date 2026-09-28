// Converts a small, safe subset of Markdown to HTML for the compose canvas
// preview. All input is HTML-escaped first, so raw HTML in the source is shown
// as text. Browser-compatible: the canvas page imports this file directly.

export function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

function safeHref(raw) {
    const href = raw.trim().replace(/&amp;/g, "&");
    if (/^(https?:|mailto:)/i.test(href)) return escapeHtml(href);
    return null;
}

function emphasis(text) {
    return text
        .replace(/\*\*([^*\n]+)\*\*/g, "<strong>$1</strong>")
        .replace(/__([^_\n]+)__/g, "<strong>$1</strong>")
        .replace(/~~([^~\n]+)~~/g, "<s>$1</s>")
        .replace(/(^|[^*\w])\*([^*\s][^*\n]*?)\*(?!\w)/g, "$1<em>$2</em>")
        .replace(/(^|[^_\w])_([^_\s][^_\n]*?)_(?!\w)/g, "$1<em>$2</em>");
}

function inline(escaped) {
    // Code spans and links become placeholders so emphasis rules never touch
    // their contents or href attributes.
    const tokens = [];
    const hold = (html) => {
        tokens.push(html);
        return `\u0000${tokens.length - 1}\u0000`;
    };
    let text = escaped.replace(/`([^`\n]+)`/g, (_, code) => hold(`<code>${code}</code>`));
    text = text.replace(/\[([^\]\n]+)\]\(([^)\s]+)\)/g, (match, label, href) => {
        const safe = safeHref(href);
        return safe ? hold(`<a href="${safe}">${emphasis(label)}</a>`) : match;
    });
    text = text.replace(/(^|[\s(])(https?:\/\/(?:(?!&quot;|&#39;|&lt;|&gt;)[^\s\u0000])+)/g, (match, lead, url) => {
        const trailing = url.match(/[.,;:!?)]+$/)?.[0] ?? "";
        const bare = trailing ? url.slice(0, -trailing.length) : url;
        return `${lead}${hold(`<a href="${bare}">${bare}</a>`)}${trailing}`;
    });
    text = emphasis(text);
    for (let depth = 0; depth < 3 && text.includes("\u0000"); depth++) {
        text = text.replace(/\u0000(\d+)\u0000/g, (_, i) => tokens[Number(i)] ?? "");
    }
    return text;
}

export function markdownToHtml(markdown) {
    const lines = String(markdown).replace(/\u0000/g, "").replace(/\r\n?/g, "\n").split("\n");
    const out = [];
    let paragraph = [];
    let list = null;
    let quote = [];

    const flushParagraph = () => {
        if (paragraph.length) out.push(`<p>${paragraph.map((l) => inline(escapeHtml(l))).join("<br>")}</p>`);
        paragraph = [];
    };
    const flushList = () => {
        if (list) out.push(`<${list.type}>${list.items.map((i) => `<li>${inline(escapeHtml(i))}</li>`).join("")}</${list.type}>`);
        list = null;
    };
    const flushQuote = () => {
        if (quote.length) out.push(`<blockquote>${markdownToHtml(quote.join("\n"))}</blockquote>`);
        quote = [];
    };
    const flushAll = () => {
        flushParagraph();
        flushList();
        flushQuote();
    };

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        const fence = line.match(/^\s*(```|~~~)/);
        if (fence) {
            flushAll();
            const body = [];
            i++;
            while (i < lines.length && !lines[i].trim().startsWith(fence[1])) body.push(lines[i++]);
            out.push(`<pre><code>${escapeHtml(body.join("\n"))}</code></pre>`);
            continue;
        }

        const quoted = line.match(/^\s*>\s?(.*)$/);
        if (quoted) {
            flushParagraph();
            flushList();
            quote.push(quoted[1]);
            continue;
        }
        flushQuote();

        if (!line.trim()) {
            flushParagraph();
            flushList();
            continue;
        }

        const heading = line.match(/^\s*(#{1,6})\s+(.*?)\s*#*\s*$/);
        if (heading) {
            flushParagraph();
            flushList();
            const level = Math.min(heading[1].length + 1, 4);
            out.push(`<h${level}>${inline(escapeHtml(heading[2]))}</h${level}>`);
            continue;
        }

        if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(line)) {
            flushParagraph();
            flushList();
            out.push("<hr>");
            continue;
        }

        const bullet = line.match(/^\s*[-*+]\s+(.*)$/);
        const numbered = line.match(/^\s*\d+[.)]\s+(.*)$/);
        if (bullet || numbered) {
            flushParagraph();
            const type = bullet ? "ul" : "ol";
            if (list && list.type !== type) flushList();
            if (!list) list = { type, items: [] };
            list.items.push((bullet || numbered)[1]);
            continue;
        }

        if (list && /^\s{2,}\S/.test(line)) {
            list.items[list.items.length - 1] += ` ${line.trim()}`;
            continue;
        }

        flushList();
        paragraph.push(line);
    }
    flushAll();
    return out.join("");
}

export function textToHtml(text) {
    return String(text)
        .replace(/\r\n?/g, "\n")
        .split("\n")
        .map((line) => escapeHtml(line))
        .join("<br>");
}

export function toTeamsHtml(message, format = "markdown") {
    return format === "text" ? textToHtml(message) : markdownToHtml(message);
}
