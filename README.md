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
| `bot.ps1` | Resident Telegram bot: commands, live alerts, scheduled digests. |
| `notify-telegram.ps1` | Send the report (or any text) to a Telegram chat. |
| `register-report.ps1` | Registers a simple twice-daily report task in Task Scheduler. |
| `register-bot.ps1` | Registers the resident bot task (at logon, auto-restart). |
| `skills/outlook-read/SKILL.md` | Agent Skill — lets an AI agent use the CLI correctly and safely. |
| `README.md` | This file. |

## Agent skill

`skills/outlook-read/SKILL.md` is an [Agent Skill](https://code.claude.com/docs/en/skills)
describing how an AI assistant (Claude Code, Codex, pi, …) should drive this CLI:
commands, flags, workflows, JSON output, and the read-only safety rules. Point your
agent's skills directory at `skills/outlook-read/` (or copy it there).

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
| `doctor` | Check the environment (PowerShell, Outlook, COM, account, inbox) |
| `whoami` | List accounts in the profile |
| `folders [-Folder N]` | List folders (or children of folder `N`) |
| `list [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d] [-Json]` | List messages |
| `unread [-N n] [-Offset k] [-Json]` | List unread messages |
| `search -Query "text" [-Folder N] [-Scope all] [-N n] [-Offset k] [-Json]` | Search subject / sender / body |
| `read -Id <entryid> [-Json]` | Full message: headers, recipients, body, links, attachment names |
| `links -Id <entryid> [-Json]` | Extract links from a message |
| `digest [-Folder N] [-N n] [-Offset k] [-Since d] [-Before d]` | Latest N with body snippets |
| `attachments -Id <entryid> [-Json]` | Attachment names / sizes / types |
| `attachment-save -Id <entryid> [-Index n] [-Dest dir] [-Json]` | Save attachment copy to a **local** temp folder |
| `calendar [-N n] [-Offset k] [-Since d] [-Before d] [-Json]` | Calendar items |
| `contacts [-N n] [-Offset k] [-Json]` | Contacts |
| `report [-Folder N] [-N n] [-Hours 12] [-State f] [-Out f] [-Update]` | Daily digest of new messages (Arabic) |
| `help` | Usage |

Flags: `-Since` / `-Before` accept `YYYY-MM-DD` or ISO. `-Scope all` searches
every mail folder. `-Json` returns machine-readable output.

## Examples

```bash
./mail doctor                     # check the environment first
./mail list -N 5
./mail list -Since 2026-09-01 -N 50
./mail unread -Json
./mail search -Query "exam schedule" -N 10
./mail search -Scope all -Query "invoice"
./mail read -Id <entryid>
./mail links -Id <entryid> -Json
./mail attachments -Id <entryid> -Json
./mail attachment-save -Id <entryid>          # -> %TEMP%\outlook-read\...
```

Example output (`list`):

```
[1] * 2026-09-19 22:12 | Course Bot <lms.smtp@example.edu> | Lecture recording GR101 [2 att] | id=0000...
[2]   2026-09-18 19:31 | Activities Team <team@example.edu> | Workshop reminder | id=0000...
```

## Scheduled daily report

`report` prints a short digest of messages newer than the last run. The last
run time is stored in a state file, so each report covers only the new window;
`-Update` advances the state.

```bash
./mail report                       # since last run (or last 12h)
./mail report -Hours 24 -N 50
./mail report -Out report.txt -Update
```

`register-report.ps1` creates a Windows Task Scheduler task that runs the
read-only report **twice a day** (default 10:00 and 18:00) and writes it to
`%LOCALAPPDATA%\outlook-read\last-report.txt`. It runs in the user's session
(only when logged on) so Classic Outlook COM is available.

```powershell
powershell -ExecutionPolicy Bypass -File register-report.ps1
powershell -ExecutionPolicy Bypass -File register-report.ps1 -Times 09:00,21:00
powershell -ExecutionPolicy Bypass -File register-report.ps1 -Remove
```

```powershell
Start-ScheduledTask -TaskName AOU-Mail-Report   # run now to test
```

This is the text that a Telegram bot (or any notifier) can send later.

## Telegram bot (commands + live alerts)

`bot.ps1` is a resident bot that:

- answers commands from Telegram,
- sends **real-time alerts** for HIGH-importance (or flagged) mail,
- sends the **digest** at the configured times (default 10:00 and 18:00),
- runs only when the user is logged on (Classic Outlook COM needs the session).

### Config (outside the repo — never commit it)

`%LOCALAPPDATA%\outlook-read\telegram.json`:

```json
{
  "token": "<bot token from @BotFather>",
  "chat_id": "<your chat id>",
  "times": ["10:00", "18:00"],
  "poll_seconds": 60,
  "alert_importance": "High",
  "alert_flagged": true,
  "digest_count": 50
}
```

Get `chat_id`: send any message to the bot, then open
`https://api.telegram.org/bot<token>/getUpdates`.

### Commands

Send `/start` once: the bot shows a **persistent keyboard of Arabic
buttons**, so normal use needs **no typing at all** — just tap. Typing still
works: ASCII names with Arabic aliases.

| Command | Arabic | Action |
|---|---|---|
| `/check` | `/جديد` | New mail since the last run |
| `/action` | `/المطلوب` | Messages that look like they need an action |
| `/exams` | `/الاختبارات` | Exam / quiz announcements |
| `/unread` | `/غير_مقروء` | Unread messages |
| `/last N` | `/آخر` | Last N messages (default 10) |
| `/read N` | `/قراءة` | Full text of message N from `/last` |
| `/search word` | `/بحث` | Search every mail folder |
| `/course CODE` | `/مقرر` | Messages about a course, e.g. `GR101` |
| `/time 10:00,18:00` | `/وقت` | Change the digest schedule |
| `/alerts off|important|all` | `/تنبيه` | Toggle alerts: off / important only / **every** new mail |
| `/interval N` | `/فحص` | How often to check mail (seconds, ≥20) |
| `/help` | `/مساعدة` | Command list |

`/action` and `/exams` are the student-focused ones: they scan recent mail for
action/deadline/exam keywords and list what matters. The engine supports the
pattern search they use via `mail search -Regex -Query "اختبار|كويز"`.

### Alert modes

| `alert_mode` | Behaviour |
|---|---|
| `off` | no alerts |
| `important` (default) | alerts only when the mail looks important (High/flagged or action/exam keywords) — sent with the **full body** |
| `all` | **every** new message gets pushed (sender + subject + short preview) |

The 🔔 التنبيهات button cycles off → important → all. `poll_seconds` (default
60) controls how fast mail is noticed; a new message is detected within one
poll interval.

### Importance score (the watchdog)

Every message gets a heuristic score **0–10** and a level, shown at the top of
every alert and in the digest:

| Score | Level |
|---|---|
| 9–10 | 🔴 عاجل |
| 7–8 | 🟠 مهم |
| 5–6 | 🟡 متوسط |
| 0–4 | 🟢 عادي |

Signals: exam/quiz (+4), deadline wording (+2), payment/fees (+2), action
keywords (+2), official AOU/Arabou sender (+2), High importance (+3), flagged
(+2), attachments (+1), unread (+1); marketing wording (−4).

The score is available in `list -Json -Snippet` as `score` + `level`, and is
used to decide whether `alert_mode = important` fires.

### Message layout

Alerts are formatted the natural way — **time first, then sender, then the
message**:

```
🕐 2026-09-20 12:42
👤 Student Announcements AOU KSA
   <std.announcements@aou.edu.sa>
📌 اعلان أداء الامتحان النصفي في فرع اخر
🚨 10/10 🔴 عاجل
────────────────
<body>
```

The body keeps its line breaks, and a divider (`──────────────`) is inserted
wherever an **Arabic run meets a Latin run**, so the two languages don't run
together.

### Images

If a message has image attachments (`.jpg/.jpeg/.png/.gif/.webp/.bmp`), up to 4
of them are uploaded to the chat with the alert (e.g. exam instruction images).
They are saved to a temp folder, uploaded, then deleted.

### Install / run

```powershell
powershell -ExecutionPolicy Bypass -File bot.ps1 -Test        # send a test batch
powershell -ExecutionPolicy Bypass -File bot.ps1 -Cmd "exams" # run one command
powershell -ExecutionPolicy Bypass -File register-bot.ps1     # resident at logon
powershell -ExecutionPolicy Bypass -File register-bot.ps1 -Remove
```

> The bot token lives only in the config file above — it is never written into
the repository.

## Read-only guarantee

`outlook-read.ps1` contains **no** `Send`, `CreateItem`, `Delete`, `Move`,
`MarkAsRead`, or `Save` calls. A write command name is rejected by the
parameter validator:

```
Cannot validate argument on parameter 'Cmd'. The argument "send"
does not belong to the set "help,doctor,whoami,folders,list,..."
```

The single exception is `attachment-save`: it copies an attachment to a
**local temp folder** for reading. It never writes to, sends from, or modifies
the mailbox. Delete the temp files when done.

## Design notes

- HTML-only messages are converted to text (`HTMLBody` → plain text) so no
  message reads as empty. Image-only messages are flagged as such.
- Internal Exchange senders expose an X500 DN in `SenderEmailAddress`; the
  tool resolves the real SMTP address via `Sender.GetExchangeUser()`
  (no object-model guard prompt in practice).
- Message metadata is exposed: **importance** (Low/Normal/High), **categories**
  and **flag** status (None/Complete/Flagged).
- **Links** are extracted from both the HTML (`href`) and the plain-text body,
  de-duplicated (`read` shows them; `links` returns them alone).
- `attachment-save` sanitizes attachment names (strips any path, replaces
  `<>:"|?*` and control chars) to prevent path traversal, and never overwrites
  an existing file.
- `search` scans the locally cached items (simple and dependency-free). For
  very large mailboxes, switch to `Restrict`/DASL.
- No network calls and no credentials stored. The only disk write is
  `attachment-save`'s temp copy.
- `doctor` validates PowerShell, Outlook process, COM, account and inbox, and
  exits non-zero on failure.

## Limitations

- Requires Classic Outlook to be installed, signed in, and launchable.
- Attachment **content** is available only by saving a copy with
  `attachment-save` (a local temp file), then reading that file.
- The `mail` wrapper needs a bash shell with `cygpath`.

## Roadmap

- [x] `attachment-save` — save an attachment to a local temp folder for reading
- [x] message importance / categories / flag status
- [x] extract links separately from the body
- [ ] render / OCR attachment content directly

## License

MIT — see [LICENSE](LICENSE).
