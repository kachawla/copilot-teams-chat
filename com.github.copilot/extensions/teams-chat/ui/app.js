import { toTeamsHtml } from "./markdown.mjs";
import { splitRecipients } from "./recipients.mjs";

const token = new URLSearchParams(location.search).get("token") || "";
const $ = (id) => document.getElementById(id);
const els = {
    account: $("account"),
    accountName: $("account-name"),
    signOut: $("sign-out"),
    signin: $("signin"),
    signInBrowser: $("sign-in-browser"),
    signInDevice: $("sign-in-device"),
    signinPending: $("signin-pending"),
    signinError: $("signin-error"),
    form: $("compose"),
    viaField: $("via-field"),
    via: $("via"),
    viaNote: $("via-note"),
    chatTarget: $("chat-target"),
    chatLabel: $("chat-label"),
    changeChat: $("change-chat"),
    toField: $("to-field"),
    to: $("to"),
    topicField: $("topic-field"),
    topic: $("topic"),
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

function recipientCount() {
    return splitRecipients(els.to.value).length;
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

function render(next) {
    // Any change to the draft (including from the agent) invalidates a pending confirmation.
    if (state && next.draft.updatedAt !== state.draft.updatedAt) hideConfirm();
    state = next;
    const { auth, draft } = state;

    const webhooks = state.webhooks || [];
    const usingWebhook = Boolean(draft.webhook);
    show(els.viaField, webhooks.length > 0 || usingWebhook);
    const options = [["", "As me, to people or a chat (needs sign-in)"], ...webhooks.map((w) => [w.name, `Teams workflow: ${w.name}${w.description ? ` (${w.description})` : ""}`])];
    if (usingWebhook && !webhooks.some((w) => w.name === draft.webhook)) options.push([draft.webhook, `Teams workflow: ${draft.webhook} (not found)`]);
    const optionsKey = JSON.stringify(options);
    if (els.via.dataset.options !== optionsKey) {
        els.via.dataset.options = optionsKey;
        els.via.replaceChildren(...options.map(([value, label]) => Object.assign(document.createElement("option"), { value, textContent: label })));
    }
    if (document.activeElement !== els.via) els.via.value = draft.webhook || "";
    show(els.viaNote, usingWebhook);

    show(els.account, auth.signedIn);
    els.accountName.textContent = auth.account
        ? `Signed in as ${auth.account.displayName || auth.account.username}${auth.account.username && auth.account.displayName ? ` (${auth.account.username})` : ""}`
        : "";
    show(els.signin, !auth.signedIn && !usingWebhook);
    const pending = auth.pending;
    if (pending?.method === "device_code") {
        els.signinPending.replaceChildren(
            "Go to ",
            Object.assign(document.createElement("a"), { href: pending.verificationUri, target: "_blank", rel: "noopener", textContent: pending.verificationUri }),
            " and enter ",
            Object.assign(document.createElement("code"), { className: "user-code", textContent: pending.userCode }),
        );
        show(els.signinPending, true);
    } else if (pending?.method === "browser") {
        els.signinPending.replaceChildren(
            "Finish signing in in your browser. ",
            Object.assign(document.createElement("a"), { href: pending.url, target: "_blank", rel: "noopener", textContent: "Open the sign-in page again" }),
        );
        show(els.signinPending, true);
    } else {
        show(els.signinPending, false);
    }
    setText(els.signinError, auth.lastError);

    const hasChat = Boolean(draft.chatId) && !usingWebhook;
    show(els.chatTarget, hasChat);
    els.chatLabel.textContent = draft.chatLabel || draft.chatId || "";
    show(els.toField, !hasChat && !usingWebhook);
    syncField(els.to, (draft.to || []).join(", "));
    syncField(els.topic, draft.topic || "");
    syncField(els.message, draft.message || "");
    if (document.activeElement !== els.format) els.format.value = draft.format || "markdown";
    show(els.topicField, !hasChat && !usingWebhook && recipientCount() > 1);

    const last = draft.lastResult;
    if (last?.ok) {
        els.result.replaceChildren(`Sent to ${last.label} at ${new Date(last.sentAt).toLocaleTimeString()}. `);
        if (last.webUrl) {
            els.result.append(Object.assign(document.createElement("a"), { href: last.webUrl, target: "_blank", rel: "noopener", textContent: "Open in Teams" }));
        }
        show(els.result, true);
    } else {
        show(els.result, false);
    }

    renderPreview();
    updateSendButton();
}

function updateSendButton() {
    const viaWebhook = Boolean(els.via.value);
    const ready =
        els.message.value.trim() && (viaWebhook || (state?.auth.signedIn && (state.draft.chatId || recipientCount() > 0)));
    els.send.disabled = busy || !ready;
}

function queueSave(patch) {
    Object.assign(pendingPatch, patch);
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushSave, 400);
    hideConfirm();
    renderPreview();
    show(els.topicField, !els.via.value && !state?.draft.chatId && recipientCount() > 1);
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

els.to.addEventListener("input", () => queueSave({ to: els.to.value }));
els.topic.addEventListener("input", () => queueSave({ topic: els.topic.value }));
els.message.addEventListener("input", () => queueSave({ message: els.message.value }));
els.format.addEventListener("change", () => queueSave({ format: els.format.value }));
els.via.addEventListener("change", () => {
    const webhook = els.via.value || null;
    const usingWebhook = Boolean(webhook);
    show(els.viaNote, usingWebhook);
    show(els.toField, !usingWebhook && !state?.draft.chatId);
    show(els.chatTarget, !usingWebhook && Boolean(state?.draft.chatId));
    show(els.signin, !usingWebhook && !state?.auth.signedIn);
    queueSave({ webhook });
});
els.changeChat.addEventListener("click", () => {
    queueSave({ chatId: null, chatLabel: null });
    show(els.chatTarget, false);
    show(els.toField, true);
    els.to.focus();
});

els.form.addEventListener("submit", (event) => {
    event.preventDefault();
    run("Checking recipients…", async () => {
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
els.signInBrowser.addEventListener("click", () => run("Opening sign-in…", async () => render(await api("/api/sign-in", { method: "browser" }))));
els.signInDevice.addEventListener("click", () => run("Requesting a code…", async () => render(await api("/api/sign-in", { method: "device_code" }))));
els.signOut.addEventListener("click", () => run("Signing out…", async () => render(await api("/api/sign-out", {}))));

const events = new EventSource(`/api/events?token=${encodeURIComponent(token)}`);
events.onmessage = (event) => {
    try {
        render(JSON.parse(event.data));
    } catch {
        // Ignore malformed frames; the next update replaces the view.
    }
};
