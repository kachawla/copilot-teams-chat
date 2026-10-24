// Checks the layout that the awesome-copilot external plugin intake requires.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const errors = [];
const manifest = JSON.parse(fs.readFileSync(path.join(root, "plugin.json"), "utf8"));
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
const allowed = new Set(["$schema", "name", "version", "description", "author", "homepage", "repository", "license", "keywords", "extensions"]);

if (manifest.$schema !== "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json") errors.push("plugin.json: unexpected $schema");
if (!/^(?!.*(?:--|\.\.))[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$/.test(manifest.name)) errors.push("plugin.json: invalid name");
if (!/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(manifest.version)) errors.push("plugin.json: version must be semver");
if (manifest.version !== pkg.version) errors.push(`version mismatch: plugin.json ${manifest.version} vs package.json ${pkg.version}`);
for (const key of Object.keys(manifest)) if (!allowed.has(key)) errors.push(`plugin.json: field "${key}" is not part of Agent Plugins v1.0.0`);
for (const kw of manifest.keywords ?? []) if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(kw)) errors.push(`plugin.json: keyword "${kw}" must be lowercase and hyphenated`);
if (!manifest.keywords?.includes("canvas")) errors.push('plugin.json: keywords must include "canvas"');
if (manifest.extensions?.["com.github.copilot"]?.logo !== "assets/preview.png") errors.push('plugin.json: extensions["com.github.copilot"].logo must be "assets/preview.png"');
if (!fs.existsSync(path.join(root, "assets", "preview.png"))) errors.push("assets/preview.png is missing");

const extRoot = path.join(root, "com.github.copilot", "extensions");
const entries = fs.existsSync(extRoot)
    ? fs.readdirSync(extRoot).filter((name) => fs.existsSync(path.join(extRoot, name, "extension.mjs")))
    : [];
if (!entries.length) errors.push("com.github.copilot/extensions/<name>/extension.mjs is missing");
if (pkg.main && !fs.existsSync(path.join(root, pkg.main))) errors.push(`package.json main "${pkg.main}" does not exist`);

const marketplacePath = path.join(root, ".github", "plugin", "marketplace.json");
if (!fs.existsSync(marketplacePath)) {
    errors.push(".github/plugin/marketplace.json is missing");
} else {
    const marketplace = JSON.parse(fs.readFileSync(marketplacePath, "utf8"));
    const entry = marketplace.plugins?.find((plugin) => plugin.name === manifest.name);
    if (!entry) {
        errors.push(`marketplace.json: no plugin named "${manifest.name}"`);
    } else {
        for (const key of ["description", "license", "homepage", "repository", "author", "keywords"]) {
            if (JSON.stringify(entry[key]) !== JSON.stringify(manifest[key])) errors.push(`marketplace.json: "${key}" does not match plugin.json`);
        }
        const repo = manifest.repository.replace(/^https:\/\/github\.com\//, "");
        if (entry.source?.source !== "github" || entry.source?.repo !== repo) errors.push(`marketplace.json: source must be { "source": "github", "repo": "${repo}" }`);
    }
}

if (errors.length) {
    for (const error of errors) console.error(`✖ ${error}`);
    process.exit(1);
}
console.log(`✔ ${manifest.name}@${manifest.version}: manifest and layout OK (extensions: ${entries.join(", ")})`);
