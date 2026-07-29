---
name: presales-timesheet-tracker
description: >-
  Tracks pre-sales activity hours, manages time entries, and produces weekly Excel timesheet reports. Use this skill whenever the user wants to log time, create a time entry, update hours, check weekly activity, build a timesheet report, query logged hours, or email a weekly report. Trigger phrases: "log hours", "add time entry", "log my time", "update my hours", "timesheet", "weekly report", "how many hours did I log", "log my activity", "pre-sales hours", "time entry for", "what did I log".
metadata:
  author: Kishor Hegde
  version: 1.0.0
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
  "created_at": "2026-07-28T14:00:00Z",
  "updated_at": "2026-07-28T14:00:00Z"
}
```

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
Set `day_of_week` from the date (e.g. "Tuesday").
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
|------|-----|----------|-------|------|----------|-------|

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
- **Duplicate entries**: Always check before saving a new entry for the same customer + date.

---

## Notes for Future Enhancements

- **Multi-day entry**: Allow logging the same activity across multiple days in one go.
- **Calendar scan**: Scan Outlook calendar for the week and suggest entries based on meetings.
- **Direct Melody integration**: When an MCP connector for Melody (SAP BTP Launchpad) is available, post entries directly instead of exporting to Excel.
- **Scheduled reminders**: When an MCP Scheduler connector is available, set up automated daily reminders at 3:30 PM and a Friday 4:00 PM report email.
