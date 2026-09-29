import { ContactError, isEmail } from "./contacts.mjs";
import { openChat, openWithOs } from "./opener.mjs";
import { buildChatLink, toTeamsText } from "./teams-link.mjs";

export function normalizeRecipients(to) {
    const list = Array.isArray(to) ? to : to ? [to] : [];
    return list
        .flatMap((entry) => String(entry).split(/[;\n]/))
        .flatMap((entry) => {
            // Commas separate only lists of addresses, so "Smith, John" stays one name.
            const parts = entry.split(",").map((p) => p.trim()).filter(Boolean);
            return parts.length > 1 && parts.every((p) => p.includes("@")) ? parts : [entry];
        })
        .map((entry) => entry.trim())
        .filter(Boolean);
}

/** Turns email addresses and saved contact names into people. Never guesses an address. */
export async function resolveRecipients(contacts, to) {
    const entries = normalizeRecipients(to);
    if (!entries.length) throw new ContactError("Say who the message is for: an email address or a saved contact name.", "no_recipients");
    const saved = await contacts.list();
    const people = [];
    const unknown = [];
    for (const entry of entries) {
        if (isEmail(entry)) {
            const contact = saved.find((c) => c.email.toLowerCase() === entry.toLowerCase());
            people.push({ name: contact?.name ?? null, email: entry });
            continue;
        }
        const contact = await contacts.find(entry);
        if (contact) people.push({ name: contact.name, email: contact.email });
        else unknown.push(entry);
    }
    if (unknown.length) {
        throw Object.assign(new ContactError(
            `No email address is known for ${unknown.map((u) => `"${u}"`).join(", ")}. ` +
                "Ask the user for the address (don't guess it), then try again with the address, or save it first with teams_chat_save_contact." +
                (saved.length ? ` Saved contacts: ${saved.map((c) => c.name).join(", ")}.` : ""),
            "unknown_recipient",
        ), { names: unknown });
    }
    return people.filter((p, i) => people.findIndex((q) => q.email.toLowerCase() === p.email.toLowerCase()) === i);
}

export function formatPerson(person) {
    return person.name ? `${person.name} <${person.email}>` : person.email;
}

/** Resolves recipients and builds the Teams links without opening anything. */
export async function prepareChat({ service, to, topic, message, format }) {
    const people = await resolveRecipients(service.contacts, to);
    const text = toTeamsText(message, format);
    const emails = people.map((p) => p.email);
    const link = (app) => buildChatLink({ emails, topic, text, tenantId: service.config.tenantId, app });
    const who = people.map(formatPerson).join(", ");
    const cleanTopic = people.length > 1 && topic ? String(topic).trim() : "";
    return {
        people,
        text,
        webLink: link(false),
        appLink: link(true),
        label: people.length === 1 ? `your chat with ${who}` : `a group chat with ${who}${cleanTopic ? ` (named "${cleanTopic}" if it's new)` : ""}`,
    };
}

export function openPrepared({ service, prepared }) {
    return openChat({
        appLink: prepared.appLink,
        webLink: prepared.webLink,
        openIn: service.config.openIn,
        openUrl: service.openUrl ?? openWithOs,
    });
}
