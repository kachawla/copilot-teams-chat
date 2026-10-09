# Changelog

Bump `version` in both `plugin.json` and `package.json` for every release, even small ones. `copilot plugin update` compares versions, so an unchanged version reports "already at latest" and gives no sign that files changed. `scripts/check-manifest.mjs` checks that the two versions match.

## 0.1.1

- Add `/teams` to open the Teams message panel.
- Fix `/teams` not opening the panel in the app.
- `/teams` failures now show as an error in the app instead of only in the log.

## 0.1.0

- Draft Teams messages with deep links: Copilot opens the chat in Teams with the message ready, and you press Send.
- Contacts, so names resolve to saved email addresses.
- Compose panel for editing drafts.
