# outlook-read

**Read-only command-line access to a Classic Outlook mailbox on Windows.**

No send, no write, no delete, no mark-read. Reading only, by design.

It talks to the Outlook desktop client through COM (the standard Windows
automation interface Microsoft provides) — so it needs **no app registration,
no OAuth, no admin consent**. It works even in locked-down tenants where
Microsoft Graph, IMAP/OAuth and third-party mail clients are blocked by an
administrator.

## Why

Many university / corporate Microsoft 365 tenants:

- block **third-party apps** (Thunderbird, any IMAP client) behind admin consent, and
- block first-party CLIs (Azure CLI, Graph CLI) with `AADSTS50105` (app assignment required).

Outlook desktop itself is already allowed (it is Microsoft's own client), so
reading the local mailbox through it is the one path that works without an
administrator. That is exactly what this tool does.

## Requirements

- Windows
- **Classic** Outlook installed and signed in (not "new Outlook")
- Windows PowerShell 5.1 (`powershell.exe`)
- For the `mail` wrapper: a bash shell (Git Bash / MSYS2) with `cygpath`

## Files

| File | Purpose |
|---|---|
| `outlook-read.ps1` | The engine (read-only). Runnable on its own. |
| `mail` | Thin bash wrapper (nice argument passing). |
| `README.md` | This file. |

## Quick start

```bash
# run the engine directly
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./outlook-read.ps1 -Cmd list -N 10

# or use the wrapper
./mail list -N 10
```

Optional: put it on your PATH.

```bash
export PATH="$HOME/tools/mail:$PATH"
```

## Commands

| Command | What it does |
|---|---|
| `whoami` | List accounts in the profile |
| `folders [-Folder N]` | List folders (or children of folder `N`) |
| `list [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d] [-Json]` | List messages |
| `unread [-N n] [-Offset k] [-Json]` | List unread messages |
| `search -Query "text" [-Folder N] [-Scope all] [-N n] [-Offset k] [-Json]` | Search subject / sender / body |
| `read -Id <entryid> [-Json]` | Full message: headers, recipients, body, attachment names |
| `digest [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d]` | Latest N with body snippets |
| `attachments -Id <entryid> [-Json]` | Attachment names / sizes / types |
| `calendar [-N n] [-Offset k] [-Since d] [-Before d] [-Json]` | Calendar items |
| `contacts [-N n] [-Offset k] [-Json]` | Contacts |
| `help` | Usage |

Flags: `-Since` / `-Before` accept `YYYY-MM-DD` or ISO. `-Scope all` searches
every mail folder. `-Json` returns machine-readable output.

## Examples

```bash
./mail list -N 5
./mail list -Since 2026-09-01 -N 50
./mail unread -Json
./mail search -Query "exam schedule" -N 10
./mail search -Scope all -Query "invoice"
./mail read -Id <entryid>
./mail attachments -Id <entryid> -Json
```

Example output (`list`):

```
[1] * 2026-09-19 22:12 | Course Bot <lms.smtp@example.edu> | Lecture recording GR101 [2 att] | id=0000...
[2]   2026-09-18 19:31 | Activities Team <team@example.edu> | Workshop reminder | id=0000...
```

## Read-only guarantee

`outlook-read.ps1` contains **no** `Send`, `CreateItem`, `Delete`, `Move`,
`Save`, or `MarkAsRead` calls. A write command name is rejected by the
parameter validator:

```
Cannot validate argument on parameter 'Cmd'. The argument "send"
does not belong to the set "help,whoami,folders,list,..."
```

## Design notes

- HTML-only messages are converted to text (`HTMLBody` → plain text) so no
  message reads as empty. Image-only messages are flagged as such.
- Internal Exchange senders expose an X500 DN in `SenderEmailAddress`; the
  tool resolves the real SMTP address via `Sender.GetExchangeUser()`
  (no object-model guard prompt in practice).
- `search` scans the locally cached items (simple and dependency-free). For
  very large mailboxes, switch to `Restrict`/DASL.
- Nothing is written to disk, no network calls, no credentials stored.

## Limitations

- Attachment **contents** are not read (only metadata). See roadmap.
- Requires Classic Outlook to be installed, signed in, and launchable.
- Message importance, categories and flags are not extracted.
- The `mail` wrapper needs a bash shell with `cygpath`.

## Roadmap

- [ ] `attachment-save` — download an attachment to a local temp folder for reading (never writes to the mailbox)
- [ ] message importance / categories / flag status
- [ ] extract links separately from the body

## License

MIT — see [LICENSE](LICENSE).
