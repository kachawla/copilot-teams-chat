import { AuthManager, AuthError, NotSignedInError } from "./auth.mjs";
import { loadConfig } from "./config.mjs";
import { GraphClient, GraphError } from "./graph.mjs";
import { TokenStore } from "./token-store.mjs";
import { RecipientError, formatPerson, personSummary } from "./send.mjs";

export async function createService({ env = process.env, fetchImpl = globalThis.fetch, openUrl, authOptions = {} } = {}) {
    const config = await loadConfig({ env });
    const store = new TokenStore(config.dataDir);
    const auth = new AuthManager({ config, store, fetchImpl, ...(openUrl ? { openUrl } : {}), ...authOptions });
    const graph = new GraphClient({ auth, baseUrl: config.graphBaseUrl, fetchImpl });
    await auth.ready;
    return { config, auth, graph };
}

/** Waits briefly for an in-progress sign-in, then returns the signed-in account or throws. */
export async function requireAccount(auth, waitForPendingMs = 0) {
    await auth.getAccessToken({ waitForPendingMs });
    const { account } = await auth.status();
    if (!account) throw new NotSignedInError();
    return account;
}

export function failure(text) {
    return { textResultForLlm: text, resultType: "failure" };
}

export function describeError(err) {
    if (err instanceof RecipientError) {
        if (!err.candidates.length) return err.message;
        const list = err.candidates.map((c) => `- ${formatPerson(c)}${c.jobTitle ? ` (${c.jobTitle})` : ""}`).join("\n");
        return `${err.message}\nCandidates:\n${list}`;
    }
    if (err instanceof NotSignedInError || err instanceof AuthError || err instanceof GraphError) return err.message;
    return `Unexpected error: ${err?.message || err}`;
}

export function formatAccount(account) {
    if (!account) return "unknown account";
    const name = account.displayName || account.username || account.id;
    return account.username && account.username !== name ? `${name} (${account.username})` : name;
}

export function describePending(pending) {
    if (!pending) return "";
    if (pending.method === "device_code") {
        return `Ask the user to open ${pending.verificationUri} and enter the code ${pending.userCode}.`;
    }
    return pending.browserOpened
        ? `The Microsoft sign-in page opened in the user's browser. If they can't find it, share this link: ${pending.url}`
        : `Could not open a browser. Ask the user to open this link to sign in: ${pending.url}`;
}

export function matchesQuery(chat, query) {
    const words = String(query).toLowerCase().split(/\s+/).filter(Boolean);
    const haystack = [chat.label, chat.topic, ...chat.members.flatMap((m) => [m.displayName, m.email])]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
    return words.every((w) => haystack.includes(w));
}

export async function findPeople(graph, query) {
    const q = String(query).trim();
    if (q.includes("@")) {
        const user = await graph.getUser(q);
        if (user) return [personSummary(user)];
    }
    return (await graph.searchUsers(q)).map(personSummary);
}
