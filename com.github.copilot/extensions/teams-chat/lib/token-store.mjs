import fs from "node:fs/promises";
import path from "node:path";
import { randomBytes } from "node:crypto";

// Persists the refresh token and account summary. The access token stays in
// memory only. The file is owner-only (0600) inside an owner-only directory.
export class TokenStore {
    constructor(dir) {
        this.dir = dir;
        this.file = path.join(dir, "auth.json");
    }

    async read() {
        try {
            const data = JSON.parse(await fs.readFile(this.file, "utf8"));
            return data && typeof data.refreshToken === "string" ? data : null;
        } catch (err) {
            if (err.code === "ENOENT") return null;
            if (err instanceof SyntaxError) return null;
            throw err;
        }
    }

    async write(data) {
        await fs.mkdir(this.dir, { recursive: true, mode: 0o700 });
        await fs.chmod(this.dir, 0o700).catch(() => {});
        const tmp = `${this.file}.${randomBytes(6).toString("hex")}.tmp`;
        await fs.writeFile(tmp, JSON.stringify(data, null, 2), { mode: 0o600 });
        await fs.rename(tmp, this.file);
    }

    async clear() {
        await fs.rm(this.file, { force: true });
    }
}
