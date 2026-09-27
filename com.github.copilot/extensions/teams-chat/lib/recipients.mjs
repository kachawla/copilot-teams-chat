// Browser-compatible: the compose canvas imports this file too.

/**
 * Splits recipient input into individual entries. Semicolons and new lines
 * always separate. Commas separate only when every piece is an email address,
 * so directory names like "Smith, John" stay whole.
 */
export function splitRecipients(value) {
    const list = Array.isArray(value) ? value : value ? [value] : [];
    return list
        .flatMap((entry) => String(entry).split(/[;\n]/))
        .flatMap((entry) => {
            const parts = entry.split(",").map((part) => part.trim()).filter(Boolean);
            return parts.length > 1 && parts.every((part) => part.includes("@")) ? parts : [entry];
        })
        .map((entry) => entry.trim())
        .filter(Boolean);
}
