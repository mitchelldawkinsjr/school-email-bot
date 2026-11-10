# School Email Bot

Google Apps Script for families: read school Gmail → extract deadlines with Gemini → family Google Calendar → Sunday HTML briefing.

Personal school/family details are **not** in the script. They live on a **Config** sheet and optional **Teachers** sheet.

## Quick install

1. Create a Google Sheet and open **Extensions → Apps Script**.
2. Paste `SchoolEmailBot.gs`, save, then run **`initializeSchoolBot`** once (authorize when prompted).
3. Fill the **Config** sheet (school domain, Gmail label, children, calendar id, notify emails).
4. Optional Advanced: fill **Teachers**, set `advanced_vocab_packets` / `advanced_math_masters` to `TRUE`.
5. In Apps Script **Project Settings → Script properties**, add `GEMINI_API_KEY`.
6. Match the Apps Script project timezone to Config `timezone`.
7. In Gmail, create a filter: school domain → your Config label (Never spam, Primary).
8. Share the family calendar with the sheet owner: **Make changes to events**.
9. From the sheet menu **School Bot**:
   - Install 30-minute schedule
   - Install Sunday briefing schedule

## Behavior highlights

- Dated items auto-publish unless Possible duplicate / Possible update / Needs date.
- Blank Gemini due dates fall back to the email sent date; rows older than ~4 weeks and outside the current month become Skipped (`Age undated Needs date rows`).
- BriefingNotes without a relevant date use the email sent date for the week filter.
- “Check the office email for important dates” is not treated as a parent action.
- Weekly briefing is forward-looking (coming school week + open parent actions).

## Config keys

| Key | Purpose |
|---|---|
| `timezone` | e.g. `America/Detroit` (must match the Apps Script project) |
| `gmail_label` | Label that holds school mail |
| `school_domain` | Only `@this.domain` senders are scanned |
| `school_name` / `school_short_name` / `school_initials` | Briefing branding |
| `family_name` / `children` | Greeting + child matching |
| `family_calendar_id` | Calendar to publish into |
| `notify_emails` | Parent notice + briefing recipients |
| `briefing_hour` | Sunday send hour (default `9`) |
| `advanced_vocab_packets` | `TRUE` to attach Vocabulary Quest PDFs |
| `advanced_math_masters` | `TRUE` to attach Math Masters PDFs |
| `vocab_children` / `math_masters_children` | Limit packets to named kids (blank = all) |

Script Properties can override `GEMINI_API_KEY`, `FAMILY_CALENDAR_ID`, and `NOTIFY_EMAILS`.

## Advanced menu

**School Bot → Advanced** opens Config / Teachers and reloads the config cache.

- **Teachers sheet**: `Teacher Email | Child | Grade` pins senders to a child.
- **Vocabulary Quest**: designed practice + separate answer-key PDFs (no LLM).
- **Math Masters**: level matrix + designed timed-fact practice sheets and answer keys; email multiples win, else Test/Level numbers map to the matrix.
- Menu item **Build Math Masters worksheet now** emails a packet immediately (uses Review rows / matrix fallback).

## Notes

- Packets use DocumentApp + DriveApp only (no Cloud Drive API enablement required).
