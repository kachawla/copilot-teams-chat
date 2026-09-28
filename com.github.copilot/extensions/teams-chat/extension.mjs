import { CanvasError, createCanvas, joinSession } from "@github/copilot-sdk/extension";
import { createComposeCanvas } from "./lib/canvas.mjs";
import { createDraftsDirResolver } from "./lib/drafts.mjs";
import { createService } from "./lib/service.mjs";
import { createTools } from "./lib/tools.mjs";

const service = await createService();

// Tools and the canvas need the session object, which only exists after
// joinSession resolves, so they reach it through this late-bound proxy.
let session;
const sessionRef = {
    get capabilities() {
        return session?.capabilities;
    },
    get ui() {
        return session?.ui;
    },
    log: (...args) => session?.log(...args),
};

const draftsDir = createDraftsDirResolver({
    getWorkspacePath: () => session?.workspacePath,
    fallbackDir: service.config.dataDir,
});

session = await joinSession({
    tools: createTools({ service, session: sessionRef }),
    canvases: [createComposeCanvas({ createCanvas, CanvasError, service, session: sessionRef, draftsDir })],
});
