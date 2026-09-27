# Teams Chat for GitHub Copilot

Send Microsoft Teams chat messages from your GitHub Copilot chats.

Ask Copilot things like "tell Alice the release build is green and link the PR" or "send my notes from this chat to myself in Teams". Copilot works out who you mean, shows you the exact message and recipient, and sends it only after you confirm.

![Teams Chat for GitHub Copilot](assets/preview.png)

## Features

- **Send to people, groups, or existing chats.** One recipient uses your 1:1 chat with them. Several recipients reuse the group chat that has exactly those members, or create one. You can also target any chat by ID, or yourself ("me" sends to your Notes to self chat).
- **Confirmation before every send.** Each message shows a preview with the resolved recipient, and nothing is sent until you approve it.
- **No guessing on names.** When a name matches more than one person, Copilot gets the candidates and asks you which one you mean.
- **Markdown formatting.** Bold, italics, code, links, lists, quotes, and code blocks are converted to Teams formatting. Raw HTML is escaped.
- **Compose canvas.** A side panel where Copilot fills in a draft and you edit it, preview the Teams formatting, and click Send. Copilot cannot send from the canvas; only your click does.
- **No dependencies.** Plain Node.js built-ins. Sign-in uses the Microsoft identity platform directly (authorization code with PKCE on a loopback redirect, or device code).

## Install

In the GitHub Copilot app or CLI:

```sh
copilot plugin install kachawla/copilot-teams-chat
```

Once listed on [awesome-copilot](https://github.com/github/awesome-copilot):

```sh
copilot plugin install teams-chat@awesome-copilot
```

Requires Node.js 22 or later and a Microsoft 365 work or school account with Teams. Personal Microsoft accounts are not supported by the Teams chat APIs.

## Use

Just ask. For example:

- "Sign in to Teams."
- "Message alice@contoso.com that the deploy finished."
- "Send a summary of this PR review to Bob and Carol in Teams, and call the chat Release sync."
- "What are my recent Teams chats?" then "Post that in the Launch plan chat."
- "Send these notes to myself in Teams."
- "Open a Teams draft to Alice so I can edit it before sending."

The first time, Copilot opens the Microsoft sign-in page in your browser. Your sign-in is saved, so you only do this once per machine. If the browser flow is blocked, ask for a device code sign-in instead.

### Tools

| Tool | What it does |
| --- | --- |
| `teams_chat_sign_in` | Starts sign-in (browser or device code) and returns right away. Other tools wait for it to finish. |
| `teams_chat_status` | Shows the signed-in account and any sign-in in progress. |
| `teams_chat_sign_out` | Signs out and deletes the saved sign-in. |
| `teams_chat_list_chats` | Lists recent chats with their IDs. Can filter by chat name or member. |
| `teams_chat_find_people` | Looks up people in your organization by name or email. |
| `teams_chat_send_message` | Sends a message after you confirm it. |

### Compose canvas

The `teams-chat-compose` canvas opens a draft panel. Copilot can prefill it (`to`, `chatId`, `topic`, `message`, `format`) and update it with the `update_draft` action. It can read your edits back with `get_draft`. Drafts are stored per Copilot session, so they survive panel reloads.

## Permissions

The extension signs in as you (delegated permissions) and requests these Microsoft Graph scopes:

| Scope | Used for |
| --- | --- |
| `ChatMessage.Send` | Sending messages in chats. |
| `Chat.ReadBasic` | Listing your chats and their members, to find and label the right chat. It cannot read message content. |
| `Chat.Create` | Starting a 1:1 or group chat when one doesn't exist yet. |
| `User.ReadBasic.All` | Resolving names and email addresses to people in your organization. |
| `offline_access`, `openid`, `profile` | Staying signed in and showing which account you're using. |

The extension never reads your messages.

By default, sign-in uses Microsoft's public **Microsoft Graph Command Line Tools** app registration (`14d82eec-204b-4c2f-b7e8-296a70dab67e`), the same client that Microsoft Graph PowerShell uses. If your organization requires admin consent for these scopes, or blocks that app, ask your admin to approve it, or use your own app registration (see below).

## Configuration

Optional settings go in `~/.copilot/teams-chat/config.json` (or `$COPILOT_HOME/teams-chat/config.json`):

```json
{
  "clientId": "00000000-0000-0000-0000-000000000000",
  "tenantId": "contoso.onmicrosoft.com",
  "confirmBeforeSend": true
}
```

| Setting | Default | Notes |
| --- | --- | --- |
| `clientId` | Microsoft Graph Command Line Tools | Your own app registration. Also settable with `TEAMS_CHAT_CLIENT_ID`. |
| `tenantId` | `organizations` | Tenant ID or domain. Also settable with `TEAMS_CHAT_TENANT_ID`. |
| `confirmBeforeSend` | `true` | Set to `false` to send without the confirmation prompt. Not recommended. |
| `authorityHost` | `https://login.microsoftonline.com` | For national clouds. |
| `graphBaseUrl` | `https://graph.microsoft.com/v1.0` | For national clouds. |

### Using your own app registration

1. In the Microsoft Entra admin center, go to **App registrations** > **New registration**.
2. Under **Authentication**, add the **Mobile and desktop applications** platform with the redirect URI `http://localhost`, and turn on **Allow public client flows** (needed for device code sign-in).
3. Under **API permissions**, add the delegated Microsoft Graph permissions listed above, and grant admin consent if your organization requires it.
4. Put the application (client) ID in `config.json` as `clientId`, then run "sign in to Teams" again.

## Privacy and security

- **Your sign-in stays on your machine.** The refresh token is saved to `~/.copilot/teams-chat/auth.json` with owner-only permissions (`0600`, in a `0700` folder). The access token is kept in memory only. Sign out to delete the file.
- **Tokens only go to Microsoft.** Graph requests are refused if the URL is not under the configured Graph base URL.
- **Every send is confirmed.** The send tool asks you to approve the exact recipient and message. If the client can't show a confirmation prompt, the tool refuses to send and suggests the compose canvas instead.
- **The compose canvas is locked to its panel.** It listens on `127.0.0.1` only, checks the `Host` header, requires a random per-panel token on every API call, and serves a strict Content Security Policy.
- **Messages are sent as you.** Recipients see them as coming from your account. Copilot is told to write in your voice and not add an "AI generated" signature unless you ask.

## Limitations

- Teams channels are not supported yet, only chats (1:1, group, and meeting chats).
- Messages are text with formatting. Attachments, images, @mentions, and Adaptive Cards are not supported yet.
- Messages can be up to about 28 KB of HTML.

## Development

```sh
npm test                       # unit and integration tests against a mock Entra ID + Graph server
node scripts/check-manifest.mjs  # checks the layout awesome-copilot requires
```

To try local changes, copy `com.github.copilot/extensions/teams-chat` into `~/.copilot/extensions/` and reload extensions.

Layout:

```
plugin.json                                   plugin manifest
assets/preview.png                            listing image
com.github.copilot/extensions/teams-chat/
  extension.mjs                               entry point: registers tools and the canvas
  lib/auth.mjs                                sign-in (PKCE loopback, device code), token refresh
  lib/graph.mjs                               Microsoft Graph client (retries, 401 refresh)
  lib/send.mjs                                recipient resolution and sending
  lib/markdown.mjs                            Markdown to Teams HTML
  lib/tools.mjs                               agent tools
  lib/canvas.mjs, lib/compose-server.mjs      compose canvas and its loopback server
  ui/                                         compose canvas page
test/                                         node:test suites and the mock cloud
```

## License

[MIT](LICENSE)
