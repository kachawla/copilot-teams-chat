import { COMPOSE_CANVAS_ID } from "./tools.mjs";

// Reusing one instance ID means /teams focuses the existing panel instead of
// opening a new one each time.
export const COMMAND_INSTANCE_ID = "teams-chat-compose-panel";

/**
 * Slash commands. `openCanvas` opens or focuses a canvas instance; `log`
 * writes to the session timeline. Tests pass fakes.
 *
 * Errors are rethrown: the SDK reports a thrown error back to the app as the
 * command's failure, which is what users see. A log line alone isn't shown.
 */
export function createCommands({ openCanvas, log }) {
    return [
        {
            name: "teams",
            description: "Open the Teams message panel to draft a message and open it in Teams.",
            handler: async () => {
                try {
                    // The app requires an object input. Input is applied as a patch,
                    // so an empty object keeps any existing draft in the panel.
                    await openCanvas({ canvasId: COMPOSE_CANVAS_ID, instanceId: COMMAND_INSTANCE_ID, input: {} });
                } catch (err) {
                    const message = `Couldn't open the Teams message panel: ${err?.message || err}`;
                    await Promise.resolve(log(message, { level: "error" })).catch(() => {});
                    throw new Error(message);
                }
            },
        },
    ];
}
