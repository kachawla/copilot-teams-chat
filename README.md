# Teams Chat for GitHub Copilot

Send Microsoft Teams messages from your GitHub Copilot chats.

Ask Copilot things like "post the release notes to the team chat in Teams" or "send me a summary of this review in Teams". Copilot writes the message, shows you exactly what will be posted and where, and sends it only after you confirm.

![Teams Chat for GitHub Copilot](assets/preview.png)

## How it works

Messages go through a **Teams workflow webhook**: a workflow you create once in Teams for each chat or channel you want to post to. There's no Microsoft sign-in and no app registration. It works even in organizations that require admin approval for third-party apps.

Messages appear in Teams from the **Workflows** bot, not from your account.

## Features

- **Confirmation before every send.** You see the destination and the full message, and nothing is posted until you approve it. If the draft changes after you confirm, you're asked again.
- **Markdown formatting.** Bold, italics, links, lists, headings, and code blocks. Messages are sent as Adaptive Cards.
- **Named destinations.** Save as many workflows as you like ("team chat", "release channel") and pick one per message. If only one is saved, it's used automatically.
- **Compose canvas.** A side panel where Copilot fills in a draft and you edit it, preview it, and click Send. Copilot cannot send from the canvas; only your click does.
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

Requires Node.js 22 or later, and Teams with the Workflows app available (it is by default in Microsoft 365).

## Set up a workflow

Do this once for each chat or channel you want to post to.

1. Create the workflow, either way:
   - In Teams, open **Apps**, search for **Workflows**, and open it. Choose **Create**, then search the templates for **webhook**.
   - Or go to [make.powerautomate.com](https://make.powerautomate.com) and open **Templates**.
2. Pick **Send webhook alerts to a chat** (or **Send webhook alerts to a channel**).
3. Name it, sign in to the Teams connection if asked, and choose where to post.
4. Copy the webhook URL it shows at the end. In Power Automate, it's also on the workflow's first step, **When a Teams webhook request is received**.
5. In Copilot, say "Save this Teams workflow as team chat: <URL>". Copilot asks you to confirm before saving it.

**Posting to just yourself:** to try it out without bothering anyone, post to a private team with only you in it, using the channel template. Or edit the workflow's **Post card in a chat or channel** step and set **Post in** to **Chat with Flow bot**; messages then land in your own chat with the Workflows app.

## Use

- "Post to team chat in Teams: the deploy finished."
- "Summarize this PR review and send it to the release channel."
- "Open a Teams draft for team chat so I can edit it before sending."
- "Which Teams workflows do I have saved?"

### Tools

| Tool | What it does |
| --- | --- |
| `teams_chat_send_message` | Posts a message to a saved workflow after you confirm it. |
| `teams_chat_add_webhook` | Saves a workflow webhook URL under a name, after you confirm it. |
| `teams_chat_list_webhooks` | Lists saved workflows by name. URLs are never shown. |
| `teams_chat_remove_webhook` | Deletes a saved workflow. |
| `teams_chat_status` | Shows saved workflows and setup help. |

### Compose canvas

The `teams-chat-compose` canvas opens a draft panel. Copilot can prefill it (`webhook`, `message`, `format`) and update it with the `update_draft` action. It can read your edits back with `get_draft`. Drafts are stored per Copilot session, so they survive panel reloads.

## Configuration

Optional settings go in `~/.copilot/teams-chat/config.json` (or `$COPILOT_HOME/teams-chat/config.json`):

| Setting | Default | Notes |
| --- | --- | --- |
| `confirmBeforeSend` | `true` | Set to `false` to send without the confirmation prompt. Not recommended. |
| `allowedWebhookHosts` | `[]` | Extra host suffixes accepted for webhook URLs, for example `["contoso.com"]`. |

## Privacy and security

- **Webhook URLs are secrets.** Anyone who has one can post to that chat. The extension stores them in `~/.copilot/teams-chat/webhooks.json` with owner-only permissions (`0600`, in a `0700` folder) and never shows them back to Copilot or the canvas. If a URL leaks, delete the workflow in Teams and create a new one.
- **Only Teams hosts.** Webhook URLs must be `https` on a Teams Workflows host (`*.logic.azure.com`, `*.powerplatform.com`, their US government equivalents, or legacy `*.webhook.office.com`). So even if instructions hidden in untrusted content get into the conversation, they can't point the extension at somewhere else to send your data.
- **New destinations are confirmed.** Saving a workflow always asks you first.
- **Every send is confirmed.** If the client can't show a confirmation prompt, the tool refuses to send and suggests the compose canvas instead. If a saved workflow's URL changes between your confirmation and the send, nothing is sent. Approvals given more than 10 minutes after the prompt appeared are ignored, so a forgotten prompt can't send later.
- **No double posts on errors.** A failed post is only retried when Teams reports throttling (HTTP 429), which means it wasn't accepted.
- **The compose canvas is locked to its panel.** It listens on `127.0.0.1` only, checks the `Host` header, requires a random per-panel token on every API call, and serves a strict Content Security Policy.

## Limitations

- Messages come from the Workflows bot, not from you.
- Each workflow posts to one chat or channel. To post somewhere else, create another workflow.
- Text with formatting only. Attachments, images, and @mentions are not supported.
- Messages can be up to about 28 KB.
- Workflows belong to the person who created them. If that person leaves the organization, the workflow stops working unless it has a co-owner.

## Development

```sh
npm test                         # unit and integration tests
node scripts/check-manifest.mjs  # checks the layout awesome-copilot requires
```

To try local changes, copy `com.github.copilot/extensions/teams-chat` into `~/.copilot/extensions/` and reload extensions.

Layout:

```
plugin.json                                   plugin manifest
assets/preview.png                            listing image
com.github.copilot/extensions/teams-chat/
  extension.mjs                               entry point: registers tools and the canvas
  lib/webhooks.mjs                            webhook storage, validation, Adaptive Cards, posting
  lib/delivery.mjs                            resolves the destination, confirms, and sends
  lib/tools.mjs                               agent tools
  lib/canvas.mjs, lib/compose-server.mjs      compose canvas and its loopback server
  lib/markdown.mjs                            Markdown preview in the canvas
  ui/                                         compose canvas page
test/                                         node:test suites
```

## License

[MIT](LICENSE)
