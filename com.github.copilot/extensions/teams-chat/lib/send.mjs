import { createHash } from "node:crypto";
import { GraphError } from "./graph.mjs";
import { MAX_MESSAGE_BYTES, toTeamsHtml } from "./markdown.mjs";
import { splitRecipients } from "./recipients.mjs";

// Teams' built-in "chat with yourself" (Notes to self). Not returned by /me/chats.
export const SELF_CHAT_ID = "48:notes";
const SELF_ALIASES = new Set(["me", "myself", "self", "note to self", "notes to self", "my notes"]);

export class RecipientError extends Error {
    constructor(message, candidates = []) {
        super(message);
        this.name = "RecipientError";
        this.code = candidates.length ? "ambiguous_recipient" : "recipient_not_found";
        this.candidates = candidates;
    }
}

export function personSummary(user) {
    return {
        id: user.id,
        displayName: user.displayName || null,
        email: user.mail || user.userPrincipalName || null,
        upn: user.userPrincipalName || null,
        jobTitle: user.jobTitle || null,
    };
}

export function formatPerson(person) {
    const name = person.displayName || person.email || person.id;
    return person.email && person.email !== name ? `${name} <${person.email}>` : name;
}

export function isSelfAlias(value) {
    return SELF_ALIASES.has(String(value).trim().toLowerCase());
}

export const normalizeRecipients = splitRecipients;

function sameAddress(person, address) {
    const q = address.toLowerCase();
    return [person.email, person.upn].some((value) => value && value.toLowerCase() === q);
}

export async function resolvePerson(graph, query) {
    const q = String(query).trim();
    if (q.includes("@")) {
        const user = await graph.getUser(q);
        if (user) return personSummary(user);
        // GET /users/{upn} misses people whose mail differs from their UPN. Search
        // is prefix-based, so only an exact address match counts.
        const matches = (await graph.searchUsers(q)).map(personSummary).filter((m) => sameAddress(m, q));
        if (matches.length === 1) return matches[0];
        if (!matches.length) throw new RecipientError(`No one in your organization has the address "${q}".`);
        throw new RecipientError(`"${q}" matches more than one person. Ask the user which one they mean.`, matches);
    }
    const matches = (await graph.searchUsers(q)).map(personSummary);
    if (matches.length === 1) return matches[0];
    const exact = matches.filter((m) => m.displayName?.toLowerCase() === q.toLowerCase());
    if (exact.length === 1) return exact[0];
    if (!matches.length) {
        throw new RecipientError(`No one in your organization matches "${q}". Try their email address.`);
    }
    throw new RecipientError(
        `"${q}" matches more than one person. Ask the user which one they mean, then retry with that person's email address.`,
        matches,
    );
}

/**
 * Works out where a message will go without changing anything in Teams.
 * Returns a plan with a human-readable `label` for the confirmation prompt.
 */
export async function planSend({ graph, account, chatId, to, topic }) {
    if (!account?.id) {
        throw new Error("The signed-in account ID is unknown. Call teams_chat_sign_out, then teams_chat_sign_in.");
    }
    const self = { kind: "self", chatId: SELF_CHAT_ID, label: "your Notes to self chat", people: [] };

    if (chatId) {
        if (chatId === SELF_CHAT_ID) return self;
        try {
            const chat = await graph.getChat(chatId, account.id);
            const label = chat.chatType === "oneOnOne" ? `1:1 chat with ${chat.label}` : `chat "${chat.label}"`;
            return { kind: "existing", chatId, label, people: [], chat };
        } catch (err) {
            if (err instanceof GraphError && err.status === 404) {
                throw new RecipientError(`Chat "${chatId}" was not found. Use teams_chat_list_chats to find the right chat ID.`);
            }
            throw err;
        }
    }

    const queries = normalizeRecipients(to);
    if (!queries.length) throw new RecipientError("Provide a chatId or at least one recipient in `to`.");

    const people = [];
    for (const q of queries) {
        if (isSelfAlias(q)) continue;
        const person = await resolvePerson(graph, q);
        if (person.id !== account.id && !people.some((p) => p.id === person.id)) people.push(person);
    }
    if (!people.length) return self;

    if (people.length === 1) {
        return { kind: "oneOnOne", people, label: `1:1 chat with ${formatPerson(people[0])}` };
    }

    const wanted = new Set([account.id, ...people.map((p) => p.id)]);
    const chats = await graph.listChats({ maxPages: 3, myUserId: account.id });
    const existing = chats.find(
        (c) =>
            c.chatType === "group" &&
            c.members.length === wanted.size &&
            c.members.every((m) => wanted.has(m.userId)),
    );
    if (existing) {
        return { kind: "existing", chatId: existing.id, people, label: `group chat "${existing.label}"`, chat: existing };
    }
    const names = people.map(formatPerson).join(", ");
    return {
        kind: "newGroup",
        people,
        topic: topic || null,
        label: `a new group chat with ${names}${topic ? ` named "${topic}"` : ""}`,
    };
}

export function renderMessage(message, format = "markdown") {
    const text = String(message ?? "");
    if (!text.trim()) throw new Error("The message is empty.");
    const html = toTeamsHtml(text, format);
    const bytes = new TextEncoder().encode(html).length;
    if (bytes > MAX_MESSAGE_BYTES) {
        throw new Error(`The message is too long for Teams (${bytes} bytes of HTML, limit ${MAX_MESSAGE_BYTES}). Shorten it or split it into several messages.`);
    }
    return html;
}

export async function executeSend({ graph, account, plan, html }) {
    let chatId = plan.chatId;
    if (plan.kind === "oneOnOne") {
        chatId = (await graph.createOneOnOneChat(account.id, plan.people[0].id)).id;
    } else if (plan.kind === "newGroup") {
        chatId = (await graph.createGroupChat(account.id, plan.people.map((p) => p.id), plan.topic)).id;
    }
    const message = await graph.sendChatMessage(chatId, html);
    return {
        chatId,
        messageId: message?.id ?? null,
        webUrl: message?.webUrl ?? null,
        createdDateTime: message?.createdDateTime ?? null,
        label: plan.label,
    };
}

/** Identifies exactly what a confirmed send will do: target and rendered content. */
export function planFingerprint(plan, html) {
    const target = {
        kind: plan.kind,
        chatId: plan.chatId ?? null,
        people: (plan.people ?? []).map((p) => p.id).sort(),
        topic: plan.topic ?? null,
    };
    return createHash("sha256").update(JSON.stringify(target)).update("\0").update(html).digest("base64url");
}

export function confirmationText(plan, message, maxChars = 1200) {
    const body = String(message);
    const preview = body.length > maxChars ? `${body.slice(0, maxChars)}\n… (${body.length - maxChars} more characters)` : body;
    return `Send this Microsoft Teams message to ${plan.label}?\n\n${preview}`;
}
