import { confirmationText, deliver, prepareSend } from "./delivery.mjs";
import { describeError, failure } from "./service.mjs";

export const COMPOSE_CANVAS_ID = "teams-chat-compose";

// The tool call times out long before an unanswered prompt goes away. An
// approval that arrives after this is ignored, so a late click can't send a
// message Copilot already reported as not sent.
export const CONFIRM_WINDOW_MS = 10 * 60_000;

const SETUP_HELP =
    'In Teams, open Apps > Workflows (or https://make.powerautomate.com > Templates), choose "Send webhook alerts to a chat" ' +
    '(or "to a channel"), pick where to post, and copy the webhook URL it shows.';

/**
 * Tool definitions. `session` supplies `capabilities`, `ui.confirm`, and `log`;
 * tests pass a fake.
 */
export function createTools({ service, session, now = Date.now, confirmWindowMs = CONFIRM_WINDOW_MS }) {
    const { config, webhooks } = service;
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
            const lines = [
                saved.length
                    ? `Saved Teams workflows: ${saved.map((w) => w.name).join(", ")}.`
                    : `No Teams workflows saved yet. ${SETUP_HELP} Then add it with teams_chat_add_webhook.`,
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
            "Post a message to Microsoft Teams through a saved Teams workflow webhook. The message appears in the chat or channel " +
            "that workflow was set up for, from the Workflows bot. If only one workflow is saved, webhook can be omitted. " +
            "The user is asked to confirm each message before it is sent. " +
            "Write the message in the user's voice; do not add a signature saying it came from Copilot unless asked.",
        parameters: {
            type: "object",
            properties: {
                message: { type: "string", description: "Message body. Markdown (bold, italics, links, lists, headings, code blocks) unless format is text." },
                webhook: { type: "string", description: "Name of the saved Teams workflow to post to." },
                format: { type: "string", enum: ["markdown", "text"], description: "How to interpret message. Default markdown." },
            },
            required: ["message"],
        },
        // The handler asks the user to confirm each message with a preview, which
        // is clearer than the generic tool permission prompt.
        skipPermission: true,
        handler: async (args = {}) => {
            try {
                const { plan, payload } = await prepareSend({ service, webhook: args.webhook, message: args.message, format: args.format });
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
                return `Sent to ${result.label}. It appears from the Workflows bot.`;
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
                name: { type: "string", description: 'Short name to refer to it, like "team chat" or "release channel".' },
                url: { type: "string", description: "The webhook URL copied from the Teams workflow." },
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
                const answer = await confirmWithUser(
                    `Save the Teams workflow webhook "${args.name}" (${host})? Messages sent to it will be posted wherever that workflow posts.`,
                );
                if (answer === "declined") return { textResultForLlm: "The user declined. The webhook was not saved.", resultType: "rejected" };
                if (answer === "expired") return expiredResult("The webhook was not saved.");
                const saved = await webhooks.add(args.name, args.url, args.description);
                return `${saved.replaced ? "Replaced" : "Saved"} the Teams workflow webhook "${saved.name}". Send to it with teams_chat_send_message and webhook: "${saved.name}".`;
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
            return saved.map((w) => `- ${w.name}${w.description ? `: ${w.description}` : ""}`).join("\n");
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

    return [status, send, addWebhook, listWebhooks, removeWebhook];
}
