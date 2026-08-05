---
name: presales-timesheet-tracker
description: >-
  Tracks pre-sales activity hours, manages time entries, and produces weekly Excel timesheet reports. Use this skill whenever the user wants to log time, create a time entry, update hours, check weekly activity, build a timesheet report, query logged hours, or email a weekly report. Trigger phrases: "log hours", "add time entry", "log my time", "update my hours", "timesheet", "weekly report", "how many hours did I log", "log my activity", "pre-sales hours", "time entry for", "what did I log".
metadata:
  author: SAP Digital Solution Advisor
  version: 1.1.0
  tags: timesheet presales time-tracking reporting productivity
---

# Pre-Sales Timesheet Tracker

Track pre-sales activity hours, manage time entries by customer, and produce weekly Excel reports for Melody timesheet submission.

---

## Data Storage

All entries are saved in `presales-log.json` in the working directory. Each entry uses this structure:

```json
{
  "id": "20260728-siemens-ag-1722160000000",
  "date": "2026-07-28",
  "day_of_week": "Tuesday",
  "customer_name": "Siemens AG",
  "hours": 2.0,
  "activity_type": "CF",
  "opportunity_number": "OPP-12345",
  "notes": "Discovery call for S/4HANA migration",
  "calendar_event_id": null,
  "created_at": "2026-07-28T14:00:00Z",
  "updated_at": "2026-07-28T14:00:00Z"
}
```

`calendar_event_id` is set to the Outlook event ID when the entry is created from a calendar scan. It is `null` for manually-created entries.

If `presales-log.json` does not exist, create it with an empty array `[]` before the first entry.

---

## Workflow 1: Create a New Entry

**Triggered by:** "log hours", "add time entry", "log my time", "I worked on...", "add entry for..."

### Step 1 — Collect fields conversationally

Ask for each field in sequence. Keep it conversational — do not dump all questions at once.

1. **Customer / Partner name** — mandatory
2. **Date** — mandatory. Accept natural language ("today", "yesterday", "last Monday"). Convert to YYYY-MM-DD. Default to today if not specified.
3. **Hours** — mandatory. Must be a positive multiple of 0.5. If not, say: "Hours must be in multiples of 0.5 — did you mean [X.0] or [X.5]?"
4. **Activity type** — mandatory. Ask: "Is this Customer Facing (CF) or Non-Customer Facing (NCF)?" Accept "CF", "NCF", "customer facing", "non-customer facing" as valid inputs.
5. **Opportunity number** — optional. Ask: "Do you have an opportunity number? (say 'skip' to leave blank)"
6. **Notes** — optional. Ask: "Any notes? For example, event name or activity performed. (say 'skip' to leave blank)"

### Step 2 — Customer name validation

Before saving, run the fuzzy name check described in the **Customer Name Matching** section below.

### Step 3 — Daily hours warning

Read `presales-log.json` and sum all existing hours for the same date. If adding the new entry would bring the daily total above 8 hours, warn:

> "This would bring your total for [date] to [X] hours, which exceeds 8 hours. Are you sure you want to log this?"

Wait for explicit confirmation before proceeding.

### Step 4 — Duplicate check

Check `presales-log.json` for an existing entry with the same `customer_name` (case-insensitive) and `date`. If one exists:

> "I already have an entry for [customer] on [date]: [X] hours ([CF/NCF]). Do you want to:
> 1. Update the existing entry
> 2. Add this as a separate new entry"

Handle the response before saving.

### Step 5 — Save

Generate the entry ID as: `YYYYMMDD-<lowercase_slug_of_customer>-<unix_timestamp_ms>`.
Set `day_of_week` from the date (e.g. "Tuesday"). Set `calendar_event_id` to `null`.
Read `presales-log.json`, append the new entry, write it back with `write_file`.
Confirm: "Logged [X] hours for [customer] on [date] ([CF/NCF])."

---

## Workflow 2: Update an Existing Entry

**Triggered by:** "update hours", "change my entry", "I made a mistake", "edit time for...", "correct hours for..."

### Step 1 — Find the entry

Ask for customer name and date if not provided. Read `presales-log.json` and find entries matching the customer (using fuzzy matching) and date.

- If multiple entries match, list them and ask which one to update.
- If no match: "I couldn't find any entries for [customer] on [date]. Would you like to create a new entry instead?"

### Step 2 — Out-of-week warning

Determine today's current calendar week (Monday to Sunday). If the entry's date falls outside that range:

> "This entry is from [date], which is outside the current week. It may already have been submitted in Melody. Do you want to:
> 1. Overwrite the existing entry
> 2. Add additional time as a separate entry"

Apply the response: overwrite updates the existing record; "add" creates a new entry.

### Step 3 — Apply changes

Ask what to change (hours, activity type, notes, opportunity number). Update the relevant fields and set `updated_at` to the current timestamp. Save. Confirm: "Updated [customer] on [date]: [new_hours] hours ([CF/NCF])."

---

## Workflow 3: Ad-hoc Query

**Triggered by:** "how many hours", "what did I log", "show entries for", "time for [customer]", "what's on my timesheet", "total hours this week"

Read `presales-log.json` and filter based on the user's request. Common patterns:

- "How many hours for [customer] this week?" → filter by customer + current week, sum hours
- "What did I log on [date]?" → filter by date
- "Show all CF entries this week" → filter by activity_type = CF + current week
- "Total hours this week" → sum all hours for Mon–Sun of current week

Display results as a Markdown table:

| Date | Day | Customer | Hours | Type | Opp. No. | Notes |
|------|-----|----------|-------|------|----------|---------|

Footer line: **Total: X.0 hrs | CF: Y.0 hrs | NCF: Z.0 hrs**

If no entries match: "No entries found for that period."

---

## Workflow 4: Generate Weekly Report (Excel)

**Triggered by:** "weekly report", "build report", "generate timesheet", "report for week of..."

### Step 1 — Determine the week

If the user names a week ("week of 21 July", "last week"), parse it. Otherwise default to the current calendar week.

- Current week = Monday 00:00 to Sunday 23:59.
- Also include Saturday and Sunday of the **previous** weekend if entries exist for those days.
- Date range: start = Monday of the target week, end = Sunday of the target week.

### Step 2 — Filter entries

Read `presales-log.json`, filter for entries within the date range, sort by date ascending then customer name ascending.

### Step 3 — Generate Excel

Run these two commands in separate terminal calls:

```bash
pip3 install openpyxl
```

```bash
python3 "<skill-disk-path>/scripts/generate_excel.py" \
  --log "presales-log.json" \
  --start "YYYY-MM-DD" \
  --end "YYYY-MM-DD" \
  --output "presales-report-YYYY-Wnn.xlsx"
```

Replace `YYYY-MM-DD` with the actual start and end dates, and `YYYY-Wnn` with the ISO week label (e.g. `2026-W31`).

### Step 4 — Confirm

Display the filtered entries as a Markdown table (same format as Workflow 3), then say:

> "Weekly report saved as `presales-report-YYYY-Wnn.xlsx` in your working directory. Would you like me to email it to you?"

---

## Workflow 5: Email Weekly Report

**Triggered by:** "email the report", "send me the report", "email my timesheet", or user replies "yes" to the prompt in Workflow 4.

If the Excel file for the requested week has not been generated yet, run Workflow 4 first.

Compose the email:

- **To:** Your own email address. Find it by calling `list_emails` and reading the `from` address of a recently sent email, or ask the user to confirm their email address.
- **Subject:** `Pre-Sales Timesheet – Week YYYY-Wnn`
- **Body:**

```
Hi,

Please find attached your pre-sales timesheet for the week of [Monday date] to [Friday date].

Summary:
- Total hours: [X.0]
- Customer Facing (CF): [Y.0] hrs
- Non-Customer Facing (NCF): [Z.0] hrs
- Entries: [N]

Please update Melody with these entries.

[Markdown table of all entries for the week]

This report was generated by Joule Pre-Sales Timesheet Tracker.
```

Note: Email sending requires Outlook send capability to be available. If email sending is not available, generate the Excel file, display the email body above, and ask the user to send it manually.

---

## Workflow 6: Calendar Scan & Auto-Log

**Triggered by:** ONLY when the user explicitly asks Joule to scan the calendar. Do NOT trigger automatically. Example phrases: "scan my calendar", "read my calendar for today", "check my Teams calendar this week", "log from my calendar", "what meetings did I have today".

### Step 1 — Determine the scan period

If the user specifies a date or range, use it. If not, ask: "Which date or week should I scan? (e.g. 'today', 'this week', 'last Monday')"

### Step 2 — Fetch calendar events

Call `list_calendar_events` for the requested date range with `limit: 50`.

Filter OUT events that are clearly not pre-sales work:
- All-day events (public holidays, out-of-office markers)
- Events with no attendees (personal reminders, blocked time)
- Events with keywords suggesting internal non-sales activities: "payroll", "labour code", "HR", "Q&A", "townhall" (from non-sales organizers), "birthday", "anniversary", "holiday", "upside tracker", "refresh"

For borderline events, include them and let the user decide.

### Step 3 — Fetch full event details & classify CF / NCF

For each remaining event, call `get_calendar_event` to retrieve the full attendee list.

Apply this classification rule:
- **CF (Customer Facing)**: at least 1 attendee has a **non-@sap.com** email address
- **NCF (Non-Customer Facing)**: ALL attendees have **@sap.com** email addresses
- **Unknown**: attendee list is empty, private, or unavailable → flag for user confirmation

IMPORTANT: Make all `get_calendar_event` calls in a **single message as parallel tool calls** — they have no dependencies on each other. This avoids fetching events one at a time.

### Step 4 — Calculate hours

Calculate each event's duration in hours from start and end times. Round to the nearest 0.5:
- Less than 15 min → 0.5h (minimum loggable unit)
- 15–44 min → 0.5h
- 45–74 min → 1.0h
- 75–104 min → 1.5h
- And so on (round to nearest 0.5)

### Step 5 — Identify the customer name

For each event, attempt to identify the customer name:
1. Parse the event title for company/customer names (e.g. "Abbott X SAP" → "Abbott", "Prep for Siemens call" → "Siemens")
2. Check external attendee email domains (e.g. someone@abbott.com → "Abbott")
3. Cross-reference against existing customer names in `presales-log.json` using fuzzy matching

If the customer name is clear → pre-fill it.
If the name is ambiguous or not identifiable → flag as "Customer: ?" and ask the user after presenting the table.
For NCF events → leave customer name as the event title or blank.

### Step 6 — Look up previous opportunity ID

For each entry where a customer name is identified:
1. Search `presales-log.json` for the most recent entry for that customer that has a non-blank `opportunity_number`.
2. If found → pre-fill and note: "Using previous Opp ID [OPP-ID] for [customer]. Confirm or change?"
3. If not found → ask: "Do you have an opportunity number for [customer]? (say 'skip' to leave blank)"

### Step 7 — Present suggested entries for confirmation

Present all identified entries in a summary table before saving anything:

| # | Date | Event | Hours | Type | Customer | Opp. No. |
|---|------|-------|-------|------|----------|----------|

Below the table, list any items flagged as Unknown type or missing customer name, and ask for those specifically.

Then say:
> "Found [N] possible entries from your calendar. Tell me which to log — e.g. 'log all', 'log 1, 2, 4', or 'skip 3'. You can also correct any field: 'change #2 customer to Deloitte'."

Also ask for notes in bulk: "Any notes to add to any of these entries? (e.g. '#2 demo prep, #4 partner review session')"

### Step 8 — Deduplication check before saving

Before saving each confirmed entry, perform TWO checks:

1. **Calendar event ID check**: Search `presales-log.json` for any entry where `calendar_event_id` matches this event's ID. If found → skip silently and note: "'[event title]' was already logged — skipped."
2. **Customer + date duplicate check**: Run the standard check from Workflow 1, Step 4.

### Step 9 — Save and notify

Save each confirmed entry with `calendar_event_id` set to the event's Outlook event ID.

After all entries are saved, display a confirmation summary and always add this notice:

> "These entries cover the meeting duration only. If you did any preparation or follow-up work for any of these meetings, log those separately — just ask me to add them."

---

## Customer Name Matching

Apply this logic every time a customer name is entered or searched.

### Step 1 — Exact match (case-insensitive)
Load all distinct `customer_name` values from `presales-log.json`. If the entered name matches one exactly (ignoring case), use the stored version.

### Step 2 — Fuzzy match
If no exact match, check for similarity:
- **Abbreviation / partial name**: "Propell" vs "PropellWorks", "SAP" vs "SAP SE"
- **Spelling error**: "Simens" → likely "Siemens"; "Micorsoft" → likely "Microsoft"
- **Extra/missing words**: "Propell Industries" vs "PropellWorks Industries"
- **Word reordering**: "Works Propell" vs "PropellWorks"

If a similar name is found, ask:

> "I found a similar customer already in your log: **[existing_name]**. Is this the same customer, or a different one?"

- **Same**: use the existing canonical name.
- **Different**: proceed with the new name as entered.

### Step 3 — No match
No existing customer found. Proceed with the name as entered. It becomes the new canonical name.

---

## Validation Rules

- **Hours**: Positive multiple of 0.5. Warn (but allow) if daily total exceeds 8 hours. Maximum 24 per single entry.
- **Date**: Valid calendar date. Future dates and weekends are allowed.
- **Activity type**: Must resolve to "CF" or "NCF". Aliases: "Customer Facing" = CF, "Non-Customer Facing" = NCF.
- **Out-of-week edits**: Warn before allowing changes to entries from previous weeks.
- **Duplicate entries**: Always check by `calendar_event_id` first (for calendar-sourced entries), then by customer name + date.

---

## Notes for Future Enhancements

- **Multi-day entry**: Allow logging the same activity across multiple days in one go.
- **Direct Melody integration**: When an MCP connector for Melody (SAP BTP Launchpad) is available, post entries directly instead of exporting to Excel.
- **Scheduled reminders**: When an MCP Scheduler connector is available, set up automated daily reminders at 3:30 PM and a Friday 4:00 PM report email.
- **Interactive confirmation UI**: When Joule Work Desktop supports clickable card actions, replace the text-based confirmation flow in Workflow 6 with per-entry confirm/skip buttons.
