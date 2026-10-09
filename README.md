# Teams Chat for GitHub Copilot

Draft Microsoft Teams messages from your GitHub Copilot chats.

Ask Copilot things like "message alice@contoso.com in Teams that the deploy finished". Copilot writes the message and opens your chat with Alice in Teams with the message ready. You read it, edit it if you like, and press **Send**.

![Teams Chat for GitHub Copilot](assets/preview.png)

## How it works

The extension uses Teams [deep links](https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/deep-link-teams), the same links Teams uses to "start a chat with" someone. It builds a link with the people and the message and opens it in the Teams app (or Teams on the web).

- **No setup.** No sign-in, app registration, admin approval, or workflow to build. If you can use Teams, it works.
- **Sent as you, only when you press Send.** Nothing is posted automatically. Teams is where you review and send.
- **Anyone in your organization.** One person opens your existing 1:1 chat with them. Several people open the group chat with exactly those people, or start a new one.

## Features

- **Message people by email or by name.** Tell Copilot once that "Alice is alice@contoso.com" and after that "message Alice" works. Copilot never guesses an email address from a name.
- **Markdown in, clean text out.** Teams shows deep-link messages as plain text, so Markdown is converted: lists become bullets, links become "label (URL)", and formatting symbols are removed.
- **Compose canvas.** A side panel where Copilot fills in a draft, you edit it and see exactly what Teams will show, then click **Open in Teams**.
- **No dependencies.** Plain Node.js built-ins.

## Install

In the GitHub Copilot app or CLI:

```sh
copilot plugin install kachawla/copilot-teams-chat
```

Once listed on [awesome-copilot](https://github.com/github/awesome-copilot):

```sh
copilot plugin install teams-chat@awesome-copilot
```

Requires Node.js 22 or later, and Microsoft Teams (the app, or Teams on the web).

## Use

- "Message alice@contoso.com in Teams: can you review my PR?"
- "Remember that Bob is bob@contoso.com." Then: "Send Bob a summary of this review in Teams."
- "Message Alice and Bob about the release, and call the chat Release sync."
- "Open a Teams draft to Alice so I can edit it here first."

Or type `/teams` to open the compose panel directly.

### Tools

| Tool | What it does |
| --- | --- |
| `teams_chat_draft_message` | Opens Teams with a chat to the people you name and your message ready to send. |
| `teams_chat_save_contact` | Remembers a person's email address under a name. |
| `teams_chat_list_contacts` | Lists saved contacts. |
| `teams_chat_remove_contact` | Forgets a contact. |

### Compose canvas

Type `/teams`, or ask Copilot for a draft, to open the `teams-chat-compose` panel. If a draft is already open, `/teams` brings it back as it was. Copilot can prefill it (`to`, `topic`, `message`, `format`) and update it with the `update_draft` action. It can read your edits back with `get_draft`. Drafts are stored per Copilot session, so they survive panel reloads.

## Configuration

Optional settings go in `~/.copilot/teams-chat/config.json` (or `$COPILOT_HOME/teams-chat/config.json`):

```json
{
  "openIn": "app",
  "tenantId": "contoso.onmicrosoft.com"
}
```

| Setting | Default | Notes |
| --- | --- | --- |
| `openIn` | `app` | `app` opens the Teams app, and falls back to Teams on the web if the app isn't installed. `web` always uses Teams on the web. |
| `tenantId` | none | Your organization's tenant ID or domain. Set it if you're signed in to several organizations in Teams and chats open in the wrong one. |

## Privacy and security

- **Nothing is sent by the extension.** It only opens Teams. The message goes out when you press Send, from your own account.
- **No credentials.** The extension never signs in and never sees a token.
- **No guessed addresses.** Names only resolve through contacts you saved. The address the chat opens with is shown in Teams before you send.
- **Contacts stay on your machine.** They're stored in `~/.copilot/teams-chat/contacts.json` with owner-only permissions.
- **The compose canvas is locked to its panel.** It listens on `127.0.0.1` only, checks the `Host` header, requires a random per-panel token on every API call, and serves a strict Content Security Policy.

## Limitations

- You press Send in Teams for each message. That's intentional; it's also what makes it work without any permissions.
- Messages are plain text. Bold, headings, and clickable link text don't carry over; URLs still become links in Teams.
- Very long messages can't fit in a link (about 8,000 characters once encoded). Split them.
- Recipients are matched by the address they sign in to Teams with (their user principal name). For most people that's their email address.
- Only people your Teams account can chat with: usually your organization, plus external people if your organization allows it.

## Development

```sh
npm test                         # unit and integration tests
node scripts/check-manifest.mjs  # checks the layout awesome-copilot requires
```

When releasing, bump `version` in `plugin.json` and `package.json` and add an entry to [CHANGELOG.md](CHANGELOG.md).

To try local changes, copy `com.github.copilot/extensions/teams-chat` into `~/.copilot/extensions/` and reload extensions.

Layout:

```
plugin.json                                   plugin manifest
assets/preview.png                            listing image
com.github.copilot/extensions/teams-chat/
  extension.mjs                               entry point: registers tools and the canvas
  lib/teams-link.mjs                          deep links and Markdown to plain text
  lib/opener.mjs                              opens links in the Teams app or on the web
  lib/compose.mjs                             resolves recipients and builds the draft
  lib/contacts.mjs                            name to email address book
  lib/tools.mjs                               agent tools
  lib/commands.mjs                            the /teams slash command
  lib/canvas.mjs, lib/compose-server.mjs      compose canvas and its loopback server
  ui/                                         compose canvas page
test/                                         node:test suites
```

## License

[MIT](LICENSE)
