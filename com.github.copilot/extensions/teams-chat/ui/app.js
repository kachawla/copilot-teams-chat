import { toTeamsHtml } from "./markdown.mjs";

const token = new URLSearchParams(location.search).get("token") || "";
const $ = (id) => document.getElementById(id);
const els = {
    setup: $("setup"),
    form: $("compose"),
    via: $("via"),
    format: $("format"),
    message: $("message"),
    preview: $("preview"),
    confirm: $("confirm"),
    confirmText: $("confirm-text"),
    confirmSend: $("confirm-send"),
    confirmCancel: $("confirm-cancel"),
    send: $("send"),
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

function hideConfirm() {
    show(els.confirm, false);
    delete els.confirm.dataset.fingerprint;
}

function renderPreview() {
    const text = els.message.value;
    els.preview.innerHTML = text.trim() ? toTeamsHtml(text, els.format.value) : "";
}

function syncField(el, value) {
    if (document.activeElement !== el && el.value !== value) el.value = value;
}

function selectedWebhook() {
    return els.via.value || null;
}

function render(next) {
    // Any change to the draft (including from the agent) invalidates a pending confirmation.
    if (state && next.draft.updatedAt !== state.draft.updatedAt) hideConfirm();
    state = next;
    const { draft } = state;
    const webhooks = state.webhooks || [];

    show(els.setup, webhooks.length === 0);
    const options = webhooks.map((w) => [w.name, `${w.name}${w.description ? ` (${w.description})` : ""}`]);
    if (draft.webhook && !webhooks.some((w) => w.name.toLowerCase() === draft.webhook.toLowerCase())) {
        options.push([draft.webhook, `${draft.webhook} (not found)`]);
    }
    if (webhooks.length > 1 && !draft.webhook) options.unshift(["", "Choose a workflow"]);
    const optionsKey = JSON.stringify(options);
    if (els.via.dataset.options !== optionsKey) {
        els.via.dataset.options = optionsKey;
        els.via.replaceChildren(...options.map(([value, label]) => Object.assign(document.createElement("option"), { value, textContent: label })));
    }
    if (document.activeElement !== els.via) {
        const match = webhooks.find((w) => w.name.toLowerCase() === (draft.webhook || "").toLowerCase());
        els.via.value = match ? match.name : draft.webhook || (webhooks.length === 1 ? webhooks[0].name : "");
    }
    els.via.disabled = options.length === 0;

    syncField(els.message, draft.message || "");
    if (document.activeElement !== els.format) els.format.value = draft.format || "markdown";

    const last = draft.lastResult;
    if (last?.ok) {
        setText(els.result, `Sent to ${last.label} at ${new Date(last.sentAt).toLocaleTimeString()}.`);
    } else {
        show(els.result, false);
    }

    renderPreview();
    updateSendButton();
}

function updateSendButton() {
    els.send.disabled = busy || !els.message.value.trim() || !selectedWebhook();
}

function queueSave(patch) {
    Object.assign(pendingPatch, patch);
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushSave, 400);
    hideConfirm();
    renderPreview();
    updateSendButton();
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

async function run(label, fn) {
    busy = true;
    updateSendButton();
    setText(els.error, "");
    els.status.textContent = label;
    try {
        await fn();
    } catch (err) {
        setText(els.error, err.message);
    } finally {
        busy = false;
        els.status.textContent = "";
        updateSendButton();
    }
}

els.message.addEventListener("input", () => queueSave({ message: els.message.value }));
els.format.addEventListener("change", () => queueSave({ format: els.format.value }));
els.via.addEventListener("change", () => queueSave({ webhook: selectedWebhook() }));

els.form.addEventListener("submit", (event) => {
    event.preventDefault();
    run("Checking…", async () => {
        // Save the shown workflow too, so a single saved workflow is used even if never picked.
        queueSave({ webhook: selectedWebhook() });
        await flushSave();
        const res = await api("/api/send", { confirm: false });
        els.confirmText.textContent = `Send this message to ${res.label}?`;
        show(els.confirm, true);
        els.confirm.dataset.fingerprint = res.fingerprint;
        els.confirmSend.focus();
    });
});

els.confirmSend.addEventListener("click", () => {
    run("Sending…", async () => {
        await flushSave();
        const fingerprint = els.confirm.dataset.fingerprint;
        if (!fingerprint) return;
        const res = await api("/api/send", { confirm: true, fingerprint });
        if (res.needsConfirm) {
            els.confirmText.textContent = `The draft changed. Send this message to ${res.label}?`;
            els.confirm.dataset.fingerprint = res.fingerprint;
            return;
        }
        hideConfirm();
        if (res.state) render(res.state);
    });
});

els.confirmCancel.addEventListener("click", hideConfirm);

const events = new EventSource(`/api/events?token=${encodeURIComponent(token)}`);
events.onmessage = (event) => {
    try {
        render(JSON.parse(event.data));
    } catch {
        // Ignore malformed frames; the next update replaces the view.
    }
};
