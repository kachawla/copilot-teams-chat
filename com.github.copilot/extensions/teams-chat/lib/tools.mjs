import { openPrepared, prepareChat } from "./compose.mjs";
import { describeError, failure } from "./service.mjs";

export const COMPOSE_CANVAS_ID = "teams-chat-compose";

/**
 * Tool definitions. `session` supplies `log`; tests pass a fake.
 */
export function createTools({ service, session }) {
    const { contacts } = service;

    const draft = {
        name: "teams_chat_draft_message",
        description:
            "Open Microsoft Teams with a message ready to send to one or more people. Nothing is sent automatically: " +
            "Teams opens the chat with the message in the compose box, and the user reviews it and presses Send. " +
            "One person opens the user's 1:1 chat with them; several people open the group chat with exactly those people, " +
            "or a new one. to takes email addresses or saved contact names. Never guess an email address from a name: use an " +
            "address the user gave, a saved contact, or ask. Teams shows the message as plain text, so Markdown syntax is removed. " +
            "Write the message in the user's voice; do not add a signature saying it came from Copilot unless asked.",
        parameters: {
            type: "object",
            properties: {
                to: {
                    type: "array",
                    items: { type: "string" },
                    description: "People to message: email addresses or saved contact names.",
                },
                message: { type: "string", description: "Message text. Markdown is converted to plain text unless format is text." },
                topic: { type: "string", description: "Name for a new group chat. Ignored for 1:1 chats and existing group chats." },
                format: { type: "string", enum: ["markdown", "text"], description: "How to read message. Default markdown." },
            },
            required: ["to", "message"],
        },
        // Opening a draft sends nothing; the user's Send click in Teams is the confirmation.
        skipPermission: true,
        handler: async (args = {}) => {
            try {
                const prepared = await prepareChat({ service, to: args.to, topic: args.topic, message: args.message, format: args.format });
                const openedIn = await openPrepared({ service, prepared });
                if (!openedIn) {
                    return `Couldn't open Teams automatically. Give the user this link to open ${prepared.label} with the message ready: ${prepared.webLink}`;
                }
                await session.log?.(`Opened a Teams draft to ${prepared.label}.`);
                return (
                    `Opened ${prepared.label} in Teams${openedIn === "web" ? " on the web" : ""} with the message ready. ` +
                    "It is NOT sent yet: tell the user to review it and press Send in Teams. " +
                    `If Teams didn't open, they can use this link: ${prepared.webLink}`
                );
            } catch (err) {
                return failure(`Teams was not opened. ${describeError(err)}`);
            }
        },
    };

    const saveContact = {
        name: "teams_chat_save_contact",
        description:
            'Remember a person\'s email address under a name, so the user can later say "message Alice" in Teams. ' +
            "Only save an address the user gave or confirmed; never guess one from a name.",
        parameters: {
            type: "object",
            properties: {
                name: { type: "string", description: 'Name to use, like "Alice" or "Alice Smith".' },
                email: { type: "string", description: "Their work email address (the one they sign in to Teams with)." },
            },
            required: ["name", "email"],
        },
        handler: async (args = {}) => {
            try {
                const saved = await contacts.save(args.name, args.email);
                return saved.replaced
                    ? `Updated ${saved.name}: ${saved.replaced.email} is now ${saved.email}.`
                    : `Saved ${saved.name} <${saved.email}>.`;
            } catch (err) {
                return failure(`The contact was not saved. ${describeError(err)}`);
            }
        },
    };

    const listContacts = {
        name: "teams_chat_list_contacts",
        description: "List saved Teams contacts (names and email addresses).",
        parameters: { type: "object", properties: {} },
        skipPermission: true,
        handler: async () => {
            const saved = await contacts.list();
            return saved.length ? saved.map((c) => `- ${c.name} <${c.email}>`).join("\n") : "No contacts saved.";
        },
    };

    const removeContact = {
        name: "teams_chat_remove_contact",
        description: "Forget a saved Teams contact.",
        parameters: {
            type: "object",
            properties: { name: { type: "string", description: "Name of the saved contact." } },
            required: ["name"],
        },
        handler: async (args = {}) =>
            (await contacts.remove(args.name)) ? `Removed the contact "${args.name}".` : failure(`No contact named "${args.name}".`),
    };

    return [draft, saveContact, listContacts, removeContact];
}
