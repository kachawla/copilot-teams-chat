import fs from "node:fs/promises";
import path from "node:path";
import { randomBytes } from "node:crypto";

// Deliberately simple: one address, no display-name tricks.
const EMAIL = /^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$/;
const NAME = /^[\p{L}\p{N}][\p{L}\p{N} .'_-]{0,63}$/u;

export class ContactError extends Error {
    constructor(message, code = "contact_error") {
        super(message);
        this.name = "ContactError";
        this.code = code;
    }
}

export function isEmail(value) {
    return EMAIL.test(String(value).trim());
}

const key = (name) => String(name).trim().toLowerCase().replace(/\s+/g, " ");

/**
 * Name to email address book, so users can say "send to Alice" once they have
 * told Copilot who Alice is. Stored owner-only in the extension data folder.
 */
export class ContactStore {
    constructor(dir) {
        this.dir = dir;
        this.file = path.join(dir, "contacts.json");
    }

    async #readAll() {
        try {
            const data = JSON.parse(await fs.readFile(this.file, "utf8"));
            return data && typeof data.contacts === "object" && data.contacts ? data.contacts : {};
        } catch (err) {
            if (err.code === "ENOENT" || err instanceof SyntaxError) return {};
            throw err;
        }
    }

    async #writeAll(contacts) {
        await fs.mkdir(this.dir, { recursive: true, mode: 0o700 });
        await fs.chmod(this.dir, 0o700).catch(() => {});
        const tmp = `${this.file}.${randomBytes(6).toString("hex")}.tmp`;
        await fs.writeFile(tmp, JSON.stringify({ contacts }, null, 2), { mode: 0o600 });
        await fs.rename(tmp, this.file);
    }

    async list() {
        const all = await this.#readAll();
        return Object.values(all).sort((a, b) => a.name.localeCompare(b.name));
    }

    async find(name) {
        return (await this.#readAll())[key(name)] ?? null;
    }

    async save(name, email) {
        const cleanName = String(name ?? "").trim().replace(/\s+/g, " ");
        const cleanEmail = String(email ?? "").trim();
        if (cleanName.includes("@")) throw new ContactError("A contact name can't be an email address.", "invalid_name");
        if (!NAME.test(cleanName)) throw new ContactError(`Invalid contact name "${cleanName}".`, "invalid_name");
        if (!isEmail(cleanEmail)) throw new ContactError(`"${cleanEmail}" is not a valid email address.`, "invalid_email");
        const all = await this.#readAll();
        const replaced = all[key(cleanName)] ?? null;
        all[key(cleanName)] = { name: cleanName, email: cleanEmail };
        await this.#writeAll(all);
        return { name: cleanName, email: cleanEmail, replaced };
    }

    async remove(name) {
        const all = await this.#readAll();
        if (!all[key(name)]) return false;
        delete all[key(name)];
        await this.#writeAll(all);
        return true;
    }
}
