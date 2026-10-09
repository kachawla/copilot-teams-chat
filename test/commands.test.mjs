import assert from "node:assert/strict";
import test from "node:test";
import { COMMAND_INSTANCE_ID, createCommands } from "../com.github.copilot/extensions/teams-chat/lib/commands.mjs";

test("/teams opens the compose panel with empty input, so an existing draft is kept", async () => {
    const opened = [];
    const [command] = createCommands({ openCanvas: async (params) => opened.push(params), log: async () => {} });
    assert.equal(command.name, "teams");
    await command.handler({ commandName: "teams", args: "" });
    await command.handler({ commandName: "teams", args: "ignored" });
    assert.deepEqual(opened, [
        { canvasId: "teams-chat-compose", instanceId: COMMAND_INSTANCE_ID, input: {} },
        { canvasId: "teams-chat-compose", instanceId: COMMAND_INSTANCE_ID, input: {} },
    ]);
});

test("/teams reports a failure to open instead of throwing", async () => {
    const logs = [];
    const [command] = createCommands({
        openCanvas: async () => {
            throw new Error("canvas_provider_unavailable");
        },
        log: async (message, options) => logs.push({ message, options }),
    });
    await command.handler({ commandName: "teams", args: "" });
    assert.deepEqual(logs, [{ message: "Couldn't open the Teams message panel: canvas_provider_unavailable", options: { level: "error" } }]);
});
