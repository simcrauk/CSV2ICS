# CSV2ICS - CSV to ICS Calendar Converter

A lightweight Windows desktop application that converts CSV and TSV files into RFC 5545 compliant iCalendar (.ics) files. Built as a single executable with no external dependencies.

## Features

- **CSV and TSV Import** - Reads CSV or tab-delimited files with automatic delimiter and header detection. RFC 4180 compliant parsing (handles quoted fields, commas in values)
- **Field Mapping** - Maps CSV columns to 25 ICS fields via dropdown selectors with live preview
- **Auto-Mapping** - Automatically matches common column names (e.g. "Start Date", "Title", "Location") to the correct ICS fields
- **Flexible Date Parsing** - Supports DD/MM/YYYY, MM/DD/YYYY, and YYYY-MM-DD formats
- **All-Day Events** - Correctly handles all-day events with exclusive DTEND per RFC 5545
- **Reminders** - Up to two VALARM reminders per event, configurable per-event via CSV or as defaults (e.g. 15 minutes, 1 hour, 1 day before)
- **Recurrence Rules** - Supports daily, weekly, fortnightly, monthly, and yearly patterns with optional interval, count, and end date (RRULE)
- **Attendees and Organizer** - Required and optional attendees with ROLE and PARTSTAT, plus ORGANIZER property
- **Duration** - Alternative to End Date/Time for specifying event length (e.g. "2 hours", "30 minutes")
- **Classification** - PUBLIC, PRIVATE, or CONFIDENTIAL visibility via ICS CLASS property
- **Per-Event Timezone** - IANA timezone per event (e.g. `America/New_York`, `Europe/Paris`) via DTSTART/DTEND TZID parameter
- **Custom X-Properties** - Any unmapped CSV column whose header starts with `X-` is automatically written as an ICS extension property
- **Export Options** - Export all events to a single .ics file, or generate separate .ics files per event
- **Apple Calendar Support** - VTIMEZONE (Europe/London with BST/GMT transitions), X-WR-RELCALID, and X-APPLE-TRAVEL-ADVISORY-BEHAVIOR for travel time alerts
- **Outlook Compatibility** - X-MICROSOFT-CDO-BUSYSTATUS automatically derived from TRANSP values
- **UUID Generation** - Each event gets a unique RFC 4122 v4 UUID via Windows CryptGenRandom

## Requirements

- Windows 10 or later (x64)
- No runtime dependencies - the application is a single standalone .exe

## Building from Source

Requires Visual Studio 2026 (or later) with the C/C++ desktop workload and Windows SDK installed.

```
build.bat
```

This compiles `csv2ics.c` and the resource file `csv2ics.rc` (application icon) into `csv2ics.exe`.

## Usage

The application uses a three-page wizard:

### Page 1 - Open CSV/TSV File

1. Click **Open CSV File** and select your CSV or TSV file (delimiter is auto-detected)
2. A preview of the data is shown in the list view
3. Choose your preferred date format (DD/MM/YYYY is the default)
4. Click **Next**

### Page 2 - Map Fields

1. Each ICS field has a dropdown populated with your CSV column headers
2. Common fields are auto-mapped where possible
3. Map at minimum: **Start Date** and **Summary** (event title)
4. Left column: core fields (dates, summary, location, status, etc.)
5. Right column: extended fields (recurrence options, attendees, duration, classification)
6. A live preview shows how the first event will look in ICS format
7. Click **Next**

### Page 3 - Export

1. Choose export mode:
   - **Single file** - all events in one .ics file (default)
   - **Separate files** - one .ics file per event, saved to a folder
2. Optionally set default reminders (applied to events that don't have a reminder column, or where the reminder field is empty)
3. Tick **Apple Calendar travel time** to include travel advisory properties
4. Click **Export** and choose a save location
5. Use **Start Over** to go back to Page 1 with a new file

### About

Access via the system menu (click the application icon in the top-left corner of the title bar, or press Alt+Space) and select **About CSV2ICS...**

## Supported Fields

The mapper recognises these 25 ICS fields (plus automatic X-property passthrough):

| Field | Description | Example CSV Values |
|-------|-------------|--------------------|
| Start Date | Event start date (required) | `29/01/2026`, `2026-01-29` |
| End Date | Event end date | `30/01/2026` |
| Summary | Event title (required) | `Team Meeting` |
| Description | Event details | `Quarterly review` |
| Location | Venue or address | `Conference Room B` |
| URL | Related link | `https://example.com` |
| Categories | Event categories | `Work, Meeting` |
| Status | Event status | `CONFIRMED`, `TENTATIVE`, `CANCELLED` |
| Transparency | Free/busy indicator | `OPAQUE`, `TRANSPARENT` |
| Priority | 0-9 (0 = undefined, 1 = highest) | `1` |
| Start Time | Time component for start | `09:00`, `9:30 AM` |
| End Time | Time component for end | `17:00` |
| All Day | Whether it's an all-day event | `Yes`, `True`, `1` |
| Reminder | Minutes before event | `15`, `60`, `1440` |
| Recurrence | Repeat rule | `Daily`, `Weekly`, `Fortnightly`, `Yearly` |
| Recurrence End Date | When recurring events stop | `31/12/2026` |
| Recurrence Count | Number of occurrences | `10`, `52` |
| Recurrence Interval | Gap between occurrences | `2` (every 2 weeks/months) |
| Classification | Visibility/access level | `PUBLIC`, `PRIVATE`, `CONFIDENTIAL` |
| Organizer | Meeting organizer email | `john@example.com` |
| Required Attendees | Required attendee emails (;-separated) | `a@ex.com; b@ex.com` |
| Optional Attendees | Optional attendee emails (;-separated) | `c@ex.com` |
| Duration | Event length (alternative to End Date) | `2 hours`, `30 minutes` |
| Timezone | IANA timezone for this event | `Europe/London`, `America/New_York` |
| X-* columns | Any unmapped column starting with `X-` | Auto-written as X-property |

## Excel Templates

Two Excel files are included to help you get started:

- **CSV2ICS_Template.xlsx** — A blank template with all 25 supported column headers, data validation dropdowns, auto-filters, and a **Notes** sheet explaining each column with accepted values and examples. Save the Events sheet as CSV to use with the app.
- **CSV2ICS_Sample.xlsx** — The same template populated with 15 sample events demonstrating timed events, all-day events, multi-day spans, recurrence patterns (daily, weekly, fortnightly, monthly, yearly) with end dates/counts/intervals, attendees and organizers, durations, classifications, various reminder durations, priorities, statuses, UK locations, and both DD/MM/YYYY and YYYY-MM-DD date formats.

## Author

**Simon Craig**

Code entirely generated by Claude (Anthropic).

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).
