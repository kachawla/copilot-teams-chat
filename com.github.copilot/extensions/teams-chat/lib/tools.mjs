import { SELF_CHAT_ID, confirmationText, executeSend, planSend, renderMessage } from "./send.mjs";
import { describeError, describePending, failure, findPeople, formatAccount, matchesQuery, requireAccount } from "./service.mjs";
import { formatPerson } from "./send.mjs";

export const COMPOSE_CANVAS_ID = "teams-chat-compose";
const SIGN_IN_WAIT_MS = 120_000;

function formatDate(value) {
    if (!value) return "no recent activity";
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? String(value) : date.toISOString().slice(0, 16).replace("T", " ") + " UTC";
}

const CHAT_TYPE = { oneOnOne: "1:1", group: "group", meeting: "meeting" };

/**
 * Tool definitions. `session` supplies `capabilities`, `ui.confirm`, and `log`;
 * tests pass a fake.
 */
export function createTools({ service, session }) {
    const { auth, graph, config } = service;

    const signIn = {
        name: "teams_chat_sign_in",
        description:
            "Sign in to Microsoft Teams (work or school account) so Copilot can send Teams chat messages for the user. " +
            "Returns immediately with instructions to show the user; sign-in finishes in the background. " +
            'Use method "device_code" if the browser flow is blocked or no local browser is available.',
        parameters: {
            type: "object",
            properties: {
                method: { type: "string", enum: ["browser", "device_code"], description: "Sign-in method. Defaults to browser." },
                force: { type: "boolean", description: "Start a new sign-in even if already signed in (to switch accounts)." },
            },
        },
        handler: async (args = {}) => {
            try {
                const status = await auth.status();
                if (status.signedIn && !args.force) {
                    return `Already signed in to Microsoft Teams as ${formatAccount(status.account)}. To switch accounts, call teams_chat_sign_in with force: true.`;
                }
                if (status.signedIn && args.force) await auth.signOut();
                const pending = await auth.startSignIn({ method: args.method || "browser", force: Boolean(args.force) });
                return `${describePending(pending)}\nSign-in completes automatically once the user finishes; other teams_chat tools wait for it. Tell the user what to do, then continue.`;
            } catch (err) {
                return failure(describeError(err));
            }
        },
    };

    const status = {
        name: "teams_chat_status",
        description: "Show whether Copilot is signed in to Microsoft Teams, as which account, and any sign-in in progress.",
        parameters: { type: "object", properties: {} },
        skipPermission: true,
        handler: async () => {
            const s = await auth.status();
            const lines = [];
            if (s.signedIn) lines.push(`Signed in to Microsoft Teams as ${formatAccount(s.account)}.`);
            else if (s.pending) lines.push(`Sign-in in progress. ${describePending(s.pending)}`);
            else lines.push("Not signed in. Call teams_chat_sign_in.");
            if (s.lastError) lines.push(`Last sign-in error: ${s.lastError}`);
            lines.push(`Client ID: ${s.clientId}; tenant: ${s.tenantId}; confirm before send: ${config.confirmBeforeSend ? "on" : "off"}.`);
            if (config.confirmBeforeSend && !session.capabilities?.ui?.elicitation) {
                lines.push(`This client can't show confirmation prompts, so teams_chat_send_message won't send. Use the "${COMPOSE_CANVAS_ID}" canvas.`);
            }
            return lines.join("\n");
        },
    };

    const signOut = {
        name: "teams_chat_sign_out",
        description: "Sign out of Microsoft Teams and delete the saved sign-in from this computer.",
        parameters: { type: "object", properties: {} },
        handler: async () => {
            await auth.signOut();
            return "Signed out of Microsoft Teams. The saved sign-in was deleted.";
        },
    };

    const listChats = {
        name: "teams_chat_list_chats",
        description:
            "List the user's recent Microsoft Teams chats (1:1, group, and meeting chats) with their chat IDs. " +
            "Use query to filter by chat name or member name/email. Use the returned chatId with teams_chat_send_message.",
        parameters: {
            type: "object",
            properties: {
                query: { type: "string", description: "Optional words to match against chat names and member names or emails." },
                limit: { type: "integer", minimum: 1, maximum: 50, description: "Maximum chats to return. Default 20." },
            },
        },
        skipPermission: true,
        handler: async (args = {}) => {
            try {
                const account = await requireAccount(auth, SIGN_IN_WAIT_MS);
                const limit = Math.min(Math.max(Number(args.limit) || 20, 1), 50);
                let chats = await graph.listChats({ maxPages: args.query ? 3 : 1, myUserId: account.id });
                if (args.query) chats = chats.filter((c) => matchesQuery(c, args.query));
                chats = chats.slice(0, limit);
                if (!chats.length) {
                    return args.query ? `No recent chats match "${args.query}".` : "No recent chats found.";
                }
                const lines = chats.map(
                    (c, i) => `${i + 1}. ${c.label} (${CHAT_TYPE[c.chatType] || c.chatType}, ${formatDate(c.lastActivity)}) chatId: ${c.id}`,
                );
                return `${lines.join("\n")}\n\nThe user's own Notes to self chat has chatId ${SELF_CHAT_ID}.`;
            } catch (err) {
                return failure(describeError(err));
            }
        },
    };

    const people = {
        name: "teams_chat_find_people",
        description: "Look up people in the user's organization by name or email, to pick the right Teams recipient.",
        parameters: {
            type: "object",
            properties: { query: { type: "string", description: "Name, partial name, or email address." } },
            required: ["query"],
        },
        skipPermission: true,
        handler: async (args = {}) => {
            try {
                await requireAccount(auth, SIGN_IN_WAIT_MS);
                const found = await findPeople(graph, args.query || "");
                if (!found.length) return `No one in the organization matches "${args.query}".`;
                return found.map((p) => `- ${formatPerson(p)}${p.jobTitle ? ` (${p.jobTitle})` : ""}`).join("\n");
            } catch (err) {
                return failure(describeError(err));
            }
        },
    };

    const send = {
        name: "teams_chat_send_message",
        description:
            "Send a Microsoft Teams chat message as the signed-in user. Give either chatId (from teams_chat_list_chats) or to " +
            '(email addresses or full names; "me" sends to the user\'s Notes to self chat). One recipient uses the 1:1 chat; ' +
            "several recipients reuse a matching group chat or create one. The user is asked to confirm each message before it is sent. " +
            "Write the message in the user's voice; do not add a signature saying it came from Copilot unless asked.",
        parameters: {
            type: "object",
            properties: {
                message: { type: "string", description: "Message body. Markdown is converted to Teams formatting unless format is text." },
                to: {
                    type: "array",
                    items: { type: "string" },
                    description: 'Recipients: email addresses (preferred) or full names. Use ["me"] for Notes to self.',
                },
                chatId: { type: "string", description: "Existing chat ID. Takes precedence over to." },
                format: { type: "string", enum: ["markdown", "text"], description: "How to interpret message. Default markdown." },
                topic: { type: "string", description: "Name for a new group chat, if one has to be created." },
            },
            required: ["message"],
        },
        // The handler asks the user to confirm each message with a preview, which
        // is clearer than the generic tool permission prompt.
        skipPermission: true,
        handler: async (args = {}) => {
            try {
                const account = await requireAccount(auth, SIGN_IN_WAIT_MS);
                const html = renderMessage(args.message, args.format);
                const plan = await planSend({ graph, account, chatId: args.chatId, to: args.to, topic: args.topic });

                if (config.confirmBeforeSend) {
                    if (!session.capabilities?.ui?.elicitation) {
                        return failure(
                            `The message was not sent: this client can't show a confirmation prompt. ` +
                                `Open the "${COMPOSE_CANVAS_ID}" canvas with this draft so the user can review and send it.`,
                        );
                    }
                    const approved = await session.ui.confirm(confirmationText(plan, args.message));
                    if (!approved) {
                        return { textResultForLlm: "The user declined. The message was not sent.", resultType: "rejected" };
                    }
                }

                const result = await executeSend({ graph, account, plan, html });
                await session.log?.(`Sent a Teams message to ${plan.label}.`);
                return [
                    `Sent to ${plan.label}.`,
                    `chatId: ${result.chatId}`,
                    result.messageId ? `messageId: ${result.messageId}` : null,
                    result.webUrl ? `Open in Teams: ${result.webUrl}` : null,
                ]
                    .filter(Boolean)
                    .join("\n");
            } catch (err) {
                return failure(`The message was not sent. ${describeError(err)}`);
            }
        },
    };

    return [signIn, status, signOut, listChats, people, send];
}
