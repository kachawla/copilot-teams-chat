import { spawn } from "node:child_process";

/**
 * Opens a URL with the OS handler. Resolves true when the handler accepted it.
 * macOS `open` and Linux `xdg-open` exit non-zero when no app handles the URL
 * (for example msteams: without Teams installed); Windows reports success
 * whenever the handler starts.
 */
export function openWithOs(url, { timeoutMs = 5000 } = {}) {
    let command;
    let args;
    if (process.platform === "darwin") {
        command = "open";
        args = [url];
    } else if (process.platform === "win32") {
        command = "rundll32";
        args = ["url.dll,FileProtocolHandler", url];
    } else {
        command = "xdg-open";
        args = [url];
    }
    return new Promise((resolve) => {
        let done = false;
        const finish = (ok) => {
            if (!done) {
                done = true;
                resolve(ok);
            }
        };
        try {
            const child = spawn(command, args, { stdio: "ignore", detached: true });
            child.once("error", () => finish(false));
            child.once("exit", (code) => finish(code === 0));
            // Some handlers stay running; if it hasn't failed by now, treat it as opened.
            setTimeout(() => {
                child.unref();
                finish(true);
            }, timeoutMs).unref();
        } catch {
            finish(false);
        }
    });
}

/**
 * Opens a chat link in the Teams app, falling back to Teams on the web when
 * the app isn't available. Returns how it was opened, or null.
 */
export async function openChat({ appLink, webLink, openIn = "app", openUrl = openWithOs }) {
    if (openIn !== "web" && (await openUrl(appLink))) return "app";
    if (await openUrl(webLink)) return "web";
    return null;
}
