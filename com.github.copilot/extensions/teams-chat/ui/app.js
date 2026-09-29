import { toTeamsText } from "./teams-link.mjs";

const token = new URLSearchParams(location.search).get("token") || "";
const $ = (id) => document.getElementById(id);
const els = {
    form: $("compose"),
    to: $("to"),
    contacts: $("contacts"),
    topicField: $("topic-field"),
    topic: $("topic"),
    format: $("format"),
    message: $("message"),
    preview: $("preview"),
    open: $("open"),
    status: $("status"),
    error: $("error"),
    result: $("result"),
};

let state = null;
let saveTimer = null;
let pendingPatch = {};
let busy = false;

async function api(path, body) {
    const response = await fetch(path, {
        method: body === undefined ? "GET" : "POST",
        headers: { "x-teams-chat-token": token, "Content-Type": "application/json" },
        body: body === undefined ? undefined : JSON.stringify(body),
    });
    const json = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(json.error || `Request failed (HTTP ${response.status})`);
    return json;
}

function show(el, visible) {
    el.hidden = !visible;
}

function setText(el, text) {
    el.textContent = text || "";
    show(el, Boolean(text));
}

function recipientCount() {
    return els.to.value.split(/[;\n]/).map((s) => s.trim()).filter(Boolean).length;
}

function renderPreview() {
    let text = "";
    try {
        text = els.message.value.trim() ? toTeamsText(els.message.value, els.format.value) : "";
    } catch {
        text = "";
    }
    els.preview.textContent = text;
}

function syncField(el, value) {
    if (document.activeElement !== el && el.value !== value) el.value = value;
}

function render(next) {
    state = next;
    const { draft } = state;
    els.contacts.replaceChildren(...(state.contacts || []).map((name) => Object.assign(document.createElement("option"), { value: name })));
    syncField(els.to, (draft.to || []).join("; "));
    syncField(els.topic, draft.topic || "");
    syncField(els.message, draft.message || "");
    if (document.activeElement !== els.format) els.format.value = draft.format || "markdown";
    show(els.topicField, recipientCount() > 1);
    renderPreview();
    updateButton();
}

function updateButton() {
    els.open.disabled = busy || !els.message.value.trim() || recipientCount() === 0;
}

function queueSave(patch) {
    Object.assign(pendingPatch, patch);
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushSave, 400);
    show(els.topicField, recipientCount() > 1);
    renderPreview();
    updateButton();
}

async function flushSave() {
    clearTimeout(saveTimer);
    const patch = pendingPatch;
    pendingPatch = {};
    if (!Object.keys(patch).length) return;
    try {
        render(await api("/api/draft", patch));
    } catch (err) {
        setText(els.error, err.message);
    }
}

els.to.addEventListener("input", () => queueSave({ to: els.to.value }));
els.topic.addEventListener("input", () => queueSave({ topic: els.topic.value }));
els.message.addEventListener("input", () => queueSave({ message: els.message.value }));
els.format.addEventListener("change", () => queueSave({ format: els.format.value }));

els.form.addEventListener("submit", async (event) => {
    event.preventDefault();
    busy = true;
    updateButton();
    setText(els.error, "");
    show(els.result, false);
    els.status.textContent = "Opening Teams…";
    try {
        await flushSave();
        const res = await api("/api/open", {});
        if (res.state) render(res.state);
        els.result.replaceChildren(
            res.openedIn
                ? `Opened ${res.label} in Teams${res.openedIn === "web" ? " on the web" : ""}. Press Send there. `
                : `Couldn't open Teams automatically. `,
            Object.assign(document.createElement("a"), { href: res.link, target: "_blank", rel: "noopener", textContent: "Open link" }),
        );
        show(els.result, true);
    } catch (err) {
        setText(els.error, err.message);
    } finally {
        busy = false;
        els.status.textContent = "";
        updateButton();
    }
});

const events = new EventSource(`/api/events?token=${encodeURIComponent(token)}`);
events.onmessage = (event) => {
    try {
        render(JSON.parse(event.data));
    } catch {
        // Ignore malformed frames; the next update replaces the view.
    }
};
