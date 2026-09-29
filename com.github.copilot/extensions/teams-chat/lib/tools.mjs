import { confirmationText, deliver, prepareSend } from "./delivery.mjs";
import { describeError, failure } from "./service.mjs";

export const COMPOSE_CANVAS_ID = "teams-chat-compose";

// The tool call times out long before an unanswered prompt goes away. An
// approval that arrives after this is ignored, so a late click can't send a
// message Copilot already reported as not sent.
export const CONFIRM_WINDOW_MS = 10 * 60_000;

const SETUP_HELP =
    'For one fixed chat or channel: in Teams, open Apps > Workflows (or https://make.powerautomate.com > Templates), choose ' +
    '"Send webhook alerts to a chat" (or "to a channel"), pick where to post, and copy the webhook URL it shows (kind "chat"). ' +
    'To send to any person by email as the user: build the workflow described under "Send to anyone" in the extension README (kind "direct").';

/**
 * Tool definitions. `session` supplies `capabilities`, `ui.confirm`, and `log`;
 * tests pass a fake.
 */
export function createTools({ service, session, now = Date.now, confirmWindowMs = CONFIRM_WINDOW_MS }) {
    const { config, webhooks, contacts } = service;
    const canConfirm = () => Boolean(session.capabilities?.ui?.elicitation);

    /** Resolves "approved", "declined", or "expired". */
    const confirmWithUser = async (message) => {
        const askedAt = now();
        const approved = await session.ui.confirm(message);
        if (!approved) return "declined";
        return now() - askedAt > confirmWindowMs ? "expired" : "approved";
    };
    const expiredResult = (what) => ({
        textResultForLlm: `${what} The confirmation was answered more than ${Math.round(confirmWindowMs / 60_000)} minutes after it was asked, so it was ignored. Ask again if the user still wants this.`,
        resultType: "rejected",
    });

    const status = {
        name: "teams_chat_status",
        description: "Show the saved Teams workflow webhooks and whether sends are confirmed.",
        parameters: { type: "object", properties: {} },
        skipPermission: true,
        handler: async () => {
            const saved = await webhooks.list();
            const direct = saved.filter((w) => w.kind === "direct");
            const lines = [
                saved.length
                    ? `Saved Teams workflows: ${saved.map((w) => `${w.name} (${w.kind === "direct" ? "send to anyone, as the user" : "one fixed chat, as the Workflows bot"})`).join(", ")}.`
                    : `No Teams workflows saved yet. ${SETUP_HELP} Then add it with teams_chat_add_webhook.`,
                direct.length ? `Saved contacts: ${(await contacts.list()).length}.` : "No send-to-anyone workflow, so messages can't be sent to specific people yet.",
                `Confirm before send: ${config.confirmBeforeSend ? "on" : "off"}.`,
            ];
            if (config.confirmBeforeSend && !canConfirm()) {
                lines.push(`This client can't show confirmation prompts, so teams_chat_send_message won't send. Use the "${COMPOSE_CANVAS_ID}" canvas.`);
            }
            return lines.join("\n");
        },
    };

    const send = {
        name: "teams_chat_send_message",
        description:
            "Send a Microsoft Teams message through a saved Teams workflow. " +
            "With to (email addresses or saved contact names), it goes through a send-to-anyone workflow as the user: one person gets it in " +
            "their 1:1 chat with the user; several people get a new group chat. Never guess an email address from a name: use an address the " +
            "user gave, a saved contact, or ask. Without to, it posts to a fixed-chat workflow's chat or channel as the Workflows bot. " +
            "If only one workflow of the needed kind is saved, webhook can be omitted. The user is asked to confirm each message before it is sent. " +
            "Write the message in the user's voice; do not add a signature saying it came from Copilot unless asked.",
        parameters: {
            type: "object",
            properties: {
                message: { type: "string", description: "Message body. Markdown (bold, italics, links, lists, headings, code blocks) unless format is text." },
                to: {
                    type: "array",
                    items: { type: "string" },
                    description: "People to send to: email addresses or saved contact names. Leave out to post to a fixed chat.",
                },
                topic: { type: "string", description: "Name for the new group chat when sending to several people." },
                webhook: { type: "string", description: "Name of the saved Teams workflow to use." },
                format: { type: "string", enum: ["markdown", "text"], description: "How to interpret message. Default markdown." },
            },
            required: ["message"],
        },
        // The handler asks the user to confirm each message with a preview, which
        // is clearer than the generic tool permission prompt.
        skipPermission: true,
        handler: async (args = {}) => {
            try {
                const { plan, payload } = await prepareSend({
                    service,
                    webhook: args.webhook,
                    to: args.to,
                    topic: args.topic,
                    message: args.message,
                    format: args.format,
                });
                if (config.confirmBeforeSend) {
                    if (!canConfirm()) {
                        return failure(
                            `The message was not sent: this client can't show a confirmation prompt. Open the "${COMPOSE_CANVAS_ID}" canvas so the user can review and send it.`,
                        );
                    }
                    const answer = await confirmWithUser(confirmationText(plan, args.message));
                    if (answer === "declined") return { textResultForLlm: "The user declined. The message was not sent.", resultType: "rejected" };
                    if (answer === "expired") return expiredResult("The message was not sent.");
                }
                const result = await deliver({ service, plan, payload });
                await session.log?.(`Sent a Teams message to ${result.label}.`);
                return result.kind === "direct"
                    ? `Handed the message for ${result.label} to the "${plan.webhookName}" workflow. Teams accepted it; if it doesn't arrive, the workflow's run history in Power Automate shows why.`
                    : `Sent to ${result.label}. It appears from the Workflows bot.`;
            } catch (err) {
                return failure(`The message was not sent. ${describeError(err)}`);
            }
        },
    };

    const addWebhook = {
        name: "teams_chat_add_webhook",
        description:
            `Save a Teams workflow webhook URL under a name so messages can be posted to it. ${SETUP_HELP} ` +
            "Only use a URL the user gave you directly. The user is asked to confirm before it is saved.",
        parameters: {
            type: "object",
            properties: {
                name: { type: "string", description: 'Short name to refer to it, like "team chat", "release channel", or "direct".' },
                url: { type: "string", description: "The webhook URL copied from the Teams workflow." },
                kind: {
                    type: "string",
                    enum: ["chat", "direct"],
                    description:
                        'chat (default): made from the "Send webhook alerts" template, posts to one fixed chat or channel. ' +
                        'direct: the send-to-anyone workflow from the README, which messages any person as the user.',
                },
                description: { type: "string", description: "Optional note about where it posts." },
            },
            required: ["name", "url"],
        },
        handler: async (args = {}) => {
            try {
                let host;
                try {
                    host = new URL(String(args.url)).hostname;
                } catch {
                    return failure("The webhook was not saved. That is not a valid URL.");
                }
                // Adding a destination is as sensitive as sending: confirm it with the user directly.
                if (!canConfirm()) {
                    return failure(`The webhook was not saved: this client can't show a confirmation prompt. Add it to ${webhooks.file} by hand.`);
                }
                const kind = args.kind === "direct" ? "direct" : "chat";
                const answer = await confirmWithUser(
                    kind === "direct"
                        ? `Save the Teams workflow webhook "${args.name}" (${host}) as a send-to-anyone workflow? Copilot will be able to send Teams messages as you to people you name. You still confirm each message.`
                        : `Save the Teams workflow webhook "${args.name}" (${host})? Messages sent to it will be posted wherever that workflow posts.`,
                );
                if (answer === "declined") return { textResultForLlm: "The user declined. The webhook was not saved.", resultType: "rejected" };
                if (answer === "expired") return expiredResult("The webhook was not saved.");
                const saved = await webhooks.add(args.name, args.url, args.description, kind);
                return kind === "direct"
                    ? `${saved.replaced ? "Replaced" : "Saved"} the send-to-anyone workflow "${saved.name}". Send with teams_chat_send_message and to: [email addresses or contact names].`
                    : `${saved.replaced ? "Replaced" : "Saved"} the Teams workflow webhook "${saved.name}". Send to it with teams_chat_send_message and webhook: "${saved.name}".`;
            } catch (err) {
                return failure(`The webhook was not saved. ${describeError(err)}`);
            }
        },
    };

    const listWebhooks = {
        name: "teams_chat_list_webhooks",
        description: "List the saved Teams workflow webhooks by name. URLs are not shown.",
        parameters: { type: "object", properties: {} },
        skipPermission: true,
        handler: async () => {
            const saved = await webhooks.list();
            if (!saved.length) return "No Teams workflow webhooks are saved. Use teams_chat_add_webhook to add one.";
            return saved
                .map((w) => `- ${w.name} (${w.kind === "direct" ? "send to anyone" : "fixed chat"})${w.description ? `: ${w.description}` : ""}`)
                .join("\n");
        },
    };

    const removeWebhook = {
        name: "teams_chat_remove_webhook",
        description: "Delete a saved Teams workflow webhook.",
        parameters: {
            type: "object",
            properties: { name: { type: "string", description: "Name of the saved webhook." } },
            required: ["name"],
        },
        handler: async (args = {}) =>
            (await webhooks.remove(args.name)) ? `Removed the Teams workflow webhook "${args.name}".` : failure(`No webhook named "${args.name}".`),
    };

    const saveContact = {
        name: "teams_chat_save_contact",
        description:
            "Remember a person's email address under a name, so the user can later say \"send to Alice\" in Teams messages. " +
            "Only save an address the user gave or confirmed; never guess one from a name.",
        parameters: {
            type: "object",
            properties: {
                name: { type: "string", description: 'Name to use, like "Alice" or "Alice Smith".' },
                email: { type: "string", description: "Their work email address." },
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

    return [status, send, addWebhook, listWebhooks, removeWebhook, saveContact, listContacts, removeContact];
}
