# fixECalendar

Convert calendar entries from Microsoft Outlook PST files to iCalendar (iCal/ICS) format.

**Built for large PST files (6GB+) with automatic deduplication using an intermediate SQLite database.**

## Features

- ✅ Extract calendar appointments from PST files
- ✅ **Handle very large PST files (6GB+) efficiently**
- ✅ **Automatic deduplication** - never import the same event twice
- ✅ **Intelligent date recovery** - recovers appointments with missing/invalid dates
- ✅ **Data quality filtering** - automatically discards corrupted entries
- ✅ **Intermediate SQLite database** for reliable processing
- ✅ Export to standard iCal (.ics) format
- ✅ **Export to CSV** for Excel/Google Sheets
- ✅ Support for both ANSI and Unicode PST formats
- ✅ Batch processing of multiple PST files
- ✅ Glob pattern support for file discovery
- ✅ Merge mode to combine multiple calendars
- ✅ Full property support (attendees, locations, descriptions, reminders)
- ✅ Handle all-day events
- ✅ Preserve timezones
- ✅ Command-line interface
- ✅ TypeScript implementation with type safety

## Why Use an Intermediate Database?

When processing **large PST files (6.5GB per file)**, fixECalendar uses a SQLite database to:

1. **Prevent Duplicates**: Automatically detects and skips duplicate calendar entries based on UID, subject, start/end times, location, and organizer
2. **Handle Large Files**: Process files without running out of memory
3. **Incremental Processing**: Add multiple PST files over time without re-importing
4. **Resume Capability**: If processing is interrupted, you can resume without starting over
5. **Fast Queries**: Quickly filter and export by date range or source file

## Intelligent Date Recovery & Data Quality

fixECalendar includes advanced data recovery and quality filtering to maximize the number of usable calendar entries:

### Date Recovery Strategies

When appointments have missing or invalid dates, fixECalendar automatically attempts recovery using:

1. **Duration-based recovery**: If only one date exists, calculates the missing date using the appointment's duration field
2. **Alternative date fields**: Falls back to other date fields like `recurrenceBase`, `attendeeCriticalChange`, `creationTime`, `clientSubmitTime`, or `messageDeliveryTime`
3. **Date reversal fix**: Automatically corrects appointments where end time is before start time
4. **Zero-duration handling**: Intelligently fixes appointments with identical start/end times:
   - **Birthdays/anniversaries/holidays**: Converted to full 24-hour all-day events (midnight to midnight)
   - **Regular appointments**: Given a standard 1-hour duration

### Data Quality Filtering

To prevent database bloat and ensure clean exports:

- **Subject validation**: Automatically discards corrupted entries with no subject (typically incomplete/corrupted PST data)
- **Keyword detection**: Automatically identifies birthdays, anniversaries, holidays, and special events for proper all-day formatting

### Impact on Import Results

**Example processing statistics:**

- **Before filtering**: 72,629 total appointments in PST files
- **After recovery & filtering**: 4,911 quality calendar entries
- **Corrupted entries removed**: ~7,356 entries with "(No Subject)" discarded
- **Date recoveries**: 341 appointments recovered with sanitized dates
- **All-day events**: 92 birthdays/anniversaries properly formatted
- **Result**: 60% reduction in database bloat while maintaining data quality

**Keywords detected for all-day events:** birthday, anniversary, holiday, easter, christmas, bank holiday, good friday, st., saint

This intelligent processing ensures you get the maximum number of valid calendar entries without cluttering your database with corrupted data.

## Installation

### Prerequisites

- Node.js 18.18.0 or higher
- npm (comes with Node.js)

### Global Installation (Recommended)

```bash
npm install -g fixecalendar
```

After installation, you can use the `fixECalendar` command globally.

### Local Installation

```bash
git clone https://github.com/ddttom/fixEcalendar.git
cd fixEcalendar
npm install
npm run build
```

## Quick Start

### 🚀 Two Ways to Export

**Option A: Direct Export (Fast)**
```bash
fixECalendar archive.pst --output calendar.ics --include-private
```

**Option B: Two-Step Workflow (Flexible)**
```bash
# Step 1: Export to CSV
npx ts-node export-to-csv.ts

# Step 2: Convert CSV to ICS
npx ts-node export-to-ical.ts
```
Result: `calendar-export.ics` ready to import! 🎉

---

### Processing Large PST Files

```bash
# Process a 6.5GB PST file (stores in database)
fixECalendar huge-archive.pst --include-private

# Process multiple large PST files (deduplicates automatically)
fixECalendar archive1.pst archive2.pst archive3.pst --include-private

# View what's in the database
fixECalendar --db-stats

# Export everything to iCal
fixECalendar --output my-calendar.ics
```

The database file (`.fixecalendar.db`) is created in your current directory and persists between runs.

## Usage

### 📋 Complete Export Workflow

There are two main workflows for exporting your calendar data:

#### Workflow 1: Direct PST to ICS (Fastest)
```bash
# Import PST files and export directly to ICS
fixECalendar archive.pst --output calendar.ics --include-private
```

#### Workflow 2: PST → Database → CSV → ICS (Most Flexible)
```bash
# Step 1: Import PST files into database
fixECalendar large-file1.pst large-file2.pst --include-private

# Step 2: Check what was imported
fixECalendar --db-stats

# Step 3: Export database to CSV (for review/editing)
npx ts-node export-to-csv.ts
# Creates: calendar-export.csv (665KB, 4,084 entries)

# Step 4: Convert CSV to ICS (for calendar import)
npx ts-node export-to-ical.ts
# Creates: calendar-export.ics (790KB, 2,490 events)
```

**Why use Workflow 2?**
- ✅ Review data in spreadsheet before final export
- ✅ Edit entries in CSV if needed
- ✅ Filter/sort data in Excel or Google Sheets
- ✅ Generate statistics and reports
- ✅ Multiple format outputs (CSV + ICS) from single source

**Visual Workflow:**
```
PST File(s) → fixECalendar → Database (.fixecalendar.db)
                                    ↓
                         export-to-csv.ts → CSV (calendar-export.csv)
                                    ↓
                        export-to-ical.ts → ICS (calendar-export.ics)
                                                      ↓
                                            Import to Calendar App
```

### Basic Workflow for Large Files

**Step 1: Import PST files into database**
```bash
# Use --include-private to get ALL appointments (recommended)
fixECalendar large-file1.pst large-file2.pst --include-private
```

**Step 2: Check what was imported**
```bash
fixECalendar --db-stats
```

**Step 3: Export to iCal format**
```bash
fixECalendar --output calendar.ics
```

**Step 4: (Optional) Export to CSV for spreadsheet analysis**
```bash
npx ts-node export-to-csv.ts
# Creates calendar-export.csv with all entries
```

**Step 5: (Optional) Convert CSV to ICS**
```bash
npx ts-node export-to-ical.ts
# Creates calendar-export.ics from CSV data
```

**Step 6: (Optional) Fix corrupted recurrence dates**
```bash
npx ts-node sanitize-recurrence-dates.ts
# Fixes Microsoft Outlook bug with UNTIL=16001231 dates
# Projects corrupted dates (year < 1900) to year 2100
```

### Single Command (Import + Export)

```bash
# Process and immediately export (include private appointments)
fixECalendar input.pst --output output.ics --include-private
```

### Batch Processing with Glob Patterns

```bash
# Process all PST files in a directory
fixECalendar "archive/*.pst"

# Process all PST files recursively
fixECalendar "**/*.pst"

# View results
fixECalendar --db-stats

# Export
fixECalendar --output combined-calendar.ics
```

### Advanced Options

**Custom Database Location**
```bash
fixECalendar input.pst --database /path/to/my-calendar.db
```

**Filter by Date Range**
```bash
# Only export events from 2024
fixECalendar --output 2024.ics \
  --start-date 2024-01-01 \
  --end-date 2024-12-31
```

**Clear Database and Start Fresh**
```bash
fixECalendar --clear-db input.pst --output fresh.ics
```

**Exclude Recurring Appointments**
```bash
fixECalendar input.pst --no-recurring
```

**Include Private Appointments (Important!)**
```bash
# By default, private/confidential appointments are excluded
# Use --include-private to include ALL appointments
fixECalendar input.pst --include-private
```

**Note:** Many Outlook users have appointments marked as private/confidential. Without `--include-private`, you may only get a fraction of your calendar entries. It's recommended to use this flag for personal calendar exports.

**Verbose Logging for Debugging**
```bash
fixECalendar huge-file.pst --verbose
```

### File Status Report

When processing multiple PST files, fixECalendar automatically generates a **File Status Report** showing problematic files:

- **Files with errors**: Unreadable, corrupted, or missing calendar folders
- **Files with zero entries**: Valid PST files but no calendar data found
- **Files with only duplicates**: All entries were already in the database

Example output:
```
=== File Status Report ===

Files with errors (2):
  - contacts.pst: No calendar folder found in PST file
  - corrupt.pst: Failed to open PST file

Files with only duplicates (1):
  - backup-copy.pst: 15987 entries were duplicates
```

This helps identify PST files that may need attention without affecting the overall import process.

## Command-Line Options

### Input/Output Options

| Option | Description |
|--------|-------------|
| `[inputs...]` | PST file(s) or glob patterns to process |
| `-o, --output <path>` | Output iCal file path |
| `-d, --output-dir <dir>` | Output directory (for multiple files) |
| `-m, --merge` | Merge all PST files into one calendar |

### Calendar Options

| Option | Description |
|--------|-------------|
| `-n, --name <name>` | Set calendar name |
| `-t, --timezone <tz>` | Set timezone (e.g., "America/New_York", "UTC") |
| `--start-date <date>` | Filter events starting from date (YYYY-MM-DD) |
| `--end-date <date>` | Filter events up to date (YYYY-MM-DD) |
| `--no-recurring` | Exclude recurring appointments |
| `--include-private` | Include private/confidential appointments |

### Database Options

| Option | Description |
|--------|-------------|
| `--database <path>` | Path to database file (default: `.fixecalendar.db`) |
| `--no-dedupe` | Disable deduplication (not recommended) |
| `--clear-db` | Clear database before processing |
| `--db-stats` | Show database statistics and exit |

### Other Options

| Option | Description |
|--------|-------------|
| `-v, --verbose` | Enable verbose logging with debug information |

## How Deduplication Works

fixECalendar generates a unique hash for each calendar entry based on:

- UID (if available from Outlook)
- Subject/Title
- Start Time
- End Time
- Location
- Organizer

If an entry with the same hash already exists in the database, it's automatically skipped. This ensures that:

- Re-processing the same PST file won't create duplicates
- Processing overlapping PST files only imports unique events
- You can safely add new PST files to your existing database

## Database Management

### View Database Statistics

```bash
fixECalendar --db-stats
```

Shows:
- Total number of entries
- Number of source files processed
- Date range of all events
- Breakdown by source file

### Clear Database

```bash
fixECalendar --clear-db
```

Removes all entries and starts fresh.

### Optimize Database

After processing many files, optimize the database:

```bash
# Database is automatically optimized after processing
# But you can also run manually:
sqlite3 .fixecalendar.db "VACUUM; ANALYZE;"
```

### Export from Existing Database

```bash
# Export everything
fixECalendar --output all-events.ics

# Export specific date range
fixECalendar --output 2024-events.ics \
  --start-date 2024-01-01 \
  --end-date 2024-12-31

# Export with custom calendar name
fixECalendar --output work-calendar.ics \
  --name "Work Calendar"
```

### Export to CSV

Export your calendar entries to CSV format for analysis in Excel, Google Sheets, or other spreadsheet applications:

```bash
# Export all entries to CSV
npx ts-node export-to-csv.ts
```

This creates `calendar-export.csv` with the following columns:
- Subject
- Start Date / Start Time
- End Date / End Time
- Location
- Description
- Organizer
- All Day (Yes/No)
- Importance
- Busy Status
- Sensitivity
- Is Recurring (Yes/No)
- Recurrence Pattern (RRULE)
- Reminder (minutes)

**Benefits of CSV export:**
- Easy filtering and sorting in spreadsheet apps
- Quick date range analysis
- Generate statistics (meetings per month, busiest days, etc.)
- Import into other calendar systems that support CSV
- Create custom reports and visualizations

**Excel Compatibility:** All fields are properly quoted for correct Excel import. This ensures that:
- Leading zeros are preserved
- Dates aren't auto-formatted incorrectly
- Multi-line descriptions are escaped as literal `\n` (newlines converted to text)
- Commas within field data don't break the CSV structure
- Each calendar entry is guaranteed to be on a single CSV row

**Description Field Formatting:** Description fields are automatically cleaned and formatted for optimal display:
- Whitespace trimmed from both ends
- Multiple consecutive line breaks (3+) reduced to 2 for better readability
- Truncated to 79 characters to prevent unwieldy text in exports
- Applied consistently across all export formats (CSV and ICS)

**All-Day Event Handling:** The CSV export properly handles all-day events (birthdays, anniversaries) by normalizing dates to remove timezone offsets, ensuring single-day events display correctly.

**Subject Standardization:** Birthday and anniversary subjects are automatically standardized to include dates in `(dd/mm/yyyy)` format. For example:
- `Trevor Simmonds's Birthday (4/6/47)` → `Trevor Simmonds's Birthday (04/06/1947)`
- `Birthday - David Hamilton (1959)` → `Birthday - David Hamilton (07/04/1959)`
- `Kevin & Sue's Anniversary (5th May 1979)` → `Kevin & Sue's Anniversary (05/05/1979)`

This ensures all birthday/anniversary dates are consistently formatted across all exports.

**Automatic Deduplication:** The CSV export includes intelligent deduplication to prevent duplicate entries:
- Duplicates are detected after date normalization and subject standardization
- Uses a unique key combining subject + start date + start time
- Reports number of duplicates removed during export
- Example output: `✓ Successfully exported 4035 unique entries (49 duplicates removed after date normalization)`

This is particularly important for recurring birthday/anniversary events that may appear multiple times in the database but should only appear once in the final export after date correction.

### Convert CSV to ICS

If you have a CSV file exported from the database, you can convert it to iCalendar (ICS) format:

```bash
# Convert calendar-export.csv to calendar-export.ics
npx ts-node export-to-ical.ts
```

This creates a full RFC 5545 compliant iCalendar file from your CSV data. The script:
- Reads from `calendar-export.csv`
- Parses all calendar entries with proper date/time handling
- Unescapes literal `\n` back to actual newlines in descriptions
- Generates `calendar-export.ics` with complete iCal structure
- Preserves all properties (attendees, recurrence, reminders, etc.)
- Handles all-day events and birthdays correctly

**Use case:** This is useful when you want to:
- Share calendar data in iCal format after reviewing/editing the CSV
- Import into calendar applications (Apple Calendar, Google Calendar, Outlook)
- Create a portable calendar file from CSV data
- Convert between formats for different use cases

### Fix Corrupted Recurrence Dates

Microsoft Outlook has a known bug where recurrence UNTIL dates are sometimes stored with year 1600 (`UNTIL=16001231`). The `sanitize-recurrence-dates.ts` script fixes these corrupted dates in your existing database:

```bash
# Fix corrupted recurrence dates in database
npx ts-node sanitize-recurrence-dates.ts
```

**What it does:**
- Scans all recurrence patterns with UNTIL dates
- Identifies corrupted dates (year < 1900)
- Projects them to year 2100 for reasonable end dates
- Reports how many entries were fixed

**Example output:**
```
Opening database...
Searching for entries with corrupted recurrence dates...
Found 72 entries with UNTIL clauses
Fixed: "U3A Genealogy Together Group" - UNTIL=16001231 → UNTIL=21001231
Fixed: "Black Bin Today" - UNTIL=16001231 → UNTIL=21001231
...
✓ Fixed 71 corrupted recurrence dates (projected to year 2100)
  1 entries had valid dates (no changes needed)
```

**When to use:**
- After importing PST files with corrupted recurrence dates
- If you see recurrence patterns with year 1600 in your calendar exports
- Before exporting to CSV/ICS to ensure clean data

**Note:** Future PST imports automatically sanitize corrupted dates (v1.2.4+), so this tool is mainly for fixing existing database entries from earlier imports.

## Supported Properties

The converter maps the following Outlook calendar properties to iCal format:

| Outlook Property | iCal Property | Notes |
|-----------------|---------------|-------|
| Subject | SUMMARY | Event title |
| Start Time | DTSTART | Includes timezone |
| End Time | DTEND | Includes timezone |
| Location | LOCATION | Meeting location |
| Body/Description | DESCRIPTION | HTML tags stripped |
| Organizer | ORGANIZER | With mailto: URI |
| Attendees | ATTENDEE | Required/Optional/Resource roles |
| All-Day Flag | DTSTART/DTEND | Date-only format |
| Importance | PRIORITY | 1=High, 5=Normal, 9=Low |
| Busy Status | STATUS/TRANSP | Free/Busy/Tentative/OOF |
| Reminder | VALARM | Minutes before event |
| Sensitivity | CLASS | PUBLIC/PRIVATE/CONFIDENTIAL |
| Recurring | RRULE | Basic support (placeholder) |

## Examples for Large Files

### Example 1: Process 3 Large PST Files (6GB each)

```bash
# Import all files into database
fixECalendar archive-2020.pst archive-2021.pst archive-2022.pst

# Check results
fixECalendar --db-stats

# Output:
# === Database Statistics ===
# Total entries: 15,437
# Source files processed: 3
# Date range: 2020-01-01 to 2022-12-31
#
# Entries by source file:
#   - archive-2020.pst: 5,234 entries
#   - archive-2021.pst: 5,112 entries (123 duplicates skipped)
#   - archive-2022.pst: 5,091 entries (89 duplicates skipped)

# Export to iCal
fixECalendar --output complete-calendar.ics
```

### Example 2: Incremental Processing

```bash
# Week 1: Process first batch
fixECalendar "batch1/*.pst"

# Week 2: Add more files (deduplication automatic)
fixECalendar "batch2/*.pst"

# Week 3: Add even more
fixECalendar "batch3/*.pst"

# Any time: Export current state
fixECalendar --output current-calendar.ics
```

### Example 3: Filter and Export by Year

```bash
# Import everything
fixECalendar "all-archives/*.pst"

# Export only 2024 events
fixECalendar --output 2024.ics \
  --start-date 2024-01-01 \
  --end-date 2024-12-31 \
  --name "2024 Calendar"

# Export only 2023 events
fixECalendar --output 2023.ics \
  --start-date 2023-01-01 \
  --end-date 2023-12-31 \
  --name "2023 Calendar"
```

### Example 4: Re-process Same File (No Duplicates)

```bash
# First run
fixECalendar important.pst
# Output: 1,000 entries added

# Second run (accidentally run again)
fixECalendar important.pst
# Output: 0 entries added, 1,000 duplicates skipped

# No duplicates created!
```

## Performance for Large Files

For 6.5GB PST files:

- **Memory Usage**: ~500MB-1GB (SQLite handles overflow to disk)
- **Processing Speed**: ~1,000-2,000 entries/second (depends on CPU)
- **Typical 6GB file**: ~30-60 minutes to process
- **Database Size**: ~100MB per 10,000 entries

### Performance Tips

1. **Use SSD**: Store the database on an SSD for faster processing
2. **Verbose Mode**: Use `--verbose` to monitor progress
3. **Batch Processing**: Process multiple files in one command for better efficiency
4. **Database Location**: Use `--database` to specify a fast storage location

## Development

### Setup

```bash
git clone https://github.com/ddttom/fixEcalendar.git
cd fixEcalendar
npm install
```

### Build

```bash
npm run build
```

### Development Mode

Run without building:

```bash
npm run dev -- input.pst --output output.ics
```

### Claude Code Commands

If you're using [Claude Code](https://claude.com/claude-code), this project includes convenient slash commands for common development workflows:

| Command | Description | Steps |
|---------|-------------|-------|
| `/build-full` | Complete pre-commit build cycle | Format → Lint fix → Test → Build |
| `/build-quick` | Fast verification | Lint → Build |
| `/build-ci` | Simulate CI pipeline locally | Clean install → Lint → Format check → Build → Test with coverage |
| `/build-test` | Test-focused workflow | Build → Test |
| `/build-release` | Pre-release verification | Clean → Install → Lint → Test → Build → Verify |
| `/add-new` | Add PST file to database | Process PST with deduplication |
| `/export-workflow` | Run complete export workflow | Database → CSV → ICS |

**Example usage:**
- Type `/build-full` in Claude Code to run the complete quality check before committing
- Type `/add-new` to add a new PST file to the database without duplicates
- Type `/export-workflow` to generate both CSV and ICS files from the database
- All build commands stop immediately on first failure

## Project Structure

```
fixEcalendar/
├── src/
│   ├── index-with-db.ts         # Main CLI with database support
│   ├── index.ts                 # Legacy CLI (no database)
│   ├── database/
│   │   └── calendar-db.ts       # SQLite database manager
│   ├── parser/
│   │   ├── pst-parser.ts        # PST file parser
│   │   ├── calendar-extractor.ts # Calendar extraction
│   │   └── types.ts             # Type definitions
│   ├── converter/
│   │   ├── ical-converter.ts    # iCal generator
│   │   ├── property-mapper.ts   # Property mapping
│   │   └── types.ts             # Type definitions
│   ├── utils/
│   │   ├── logger.ts            # Logging
│   │   ├── validators.ts        # Validation
│   │   ├── error-handler.ts    # Error handling
│   │   └── text-formatter.ts   # Text formatting utilities
│   └── config/
│       └── constants.ts         # Constants
├── .fixecalendar.db            # SQLite database (generated)
├── dist/                        # Compiled JavaScript
├── tests/                       # Test files
├── package.json
├── tsconfig.json
└── README.md
```

## Recurrence Pattern Support

✅ **Full RRULE Support** - fixECalendar now has comprehensive recurrence pattern extraction using pst-extractor's RecurrencePattern class:

- **Frequencies**: DAILY, WEEKLY, MONTHLY, YEARLY
- **By Day**: Specific weekdays (BYDAY=MO,TU,WE,TH,FR)
- **By Month Day**: Day of month (BYMONTHDAY=15)
- **By Day with Ordinals**: 2nd Monday (BYDAY=2MO), Last Friday (BYDAY=-1FR)
- **Intervals**: Every N days/weeks/months (INTERVAL=2)
- **Count**: Number of occurrences (COUNT=10)
- **Until**: End date (UNTIL=20251231)

**Real examples from extracted data:**
- `FREQ=WEEKLY;BYDAY=MO;COUNT=7` - Every Monday, 7 occurrences
- `FREQ=MONTHLY;BYDAY=2MO;COUNT=2` - 2nd Monday of month, 2 times
- `FREQ=MONTHLY;BYDAY=-1FR` - Last Friday of month
- `FREQ=YEARLY;BYMONTHDAY=29` - Yearly on 29th (birthdays)

**Google Calendar Compatibility:** Version 1.2.2+ correctly formats yearly recurrence rules for RFC 5545 compliance. Outlook stores yearly events as 12-month intervals, which are now automatically converted to proper yearly intervals for compatibility with Google Calendar and other standards-compliant applications.

## Limitations

- **Attachments**: Not included in iCal export
- **Categories**: Not currently mapped
- **Private Appointments**: Excluded by default - **IMPORTANT:** Use `--include-private` to include personal/private/confidential appointments. Many users have most of their appointments marked as private, so this flag is highly recommended for personal calendar exports.

## Troubleshooting

### Google Calendar import fails silently (FIXED in v1.2.2)

**Issue:** ICS file appears valid but Google Calendar doesn't import any events, or imports fail silently.

**Cause:** Invalid RRULE format with `INTERVAL=12` for yearly recurring events. Outlook stores yearly recurrence as 12-month intervals, but RFC 5545 (iCalendar standard) expects yearly events to omit the interval or use `INTERVAL=1`.

**Solution:** Fixed in version 1.2.2+. The tool now:
- Automatically converts Outlook's 12-month intervals to proper yearly format
- Cleans up existing CSV data during conversion to ICS
- Generates RFC 5545 compliant RRULE for all yearly events

If you exported ICS files before v1.2.2, simply re-run the export:
```bash
# Regenerate from CSV
npx ts-node export-to-ical.ts

# Or regenerate entire workflow
npx ts-node export-to-csv.ts
npx ts-node export-to-ical.ts
```

The new ICS file will import successfully into Google Calendar, Apple Calendar, and all RFC 5545 compliant applications.

### "No calendar folder found in PST file" (FIXED in v1.2.3)

**Issue:** PST file contains calendar data but the tool reports "No calendar folder found".

**Cause:** The tool now uses Microsoft's PR_CONTAINER_CLASS property to detect calendar folders, which is more reliable than name-based detection. However, some PST files may have nested folder structures or corrupted folder properties.

**Solution (Fixed in v1.2.3):**
- The tool now searches ALL folders in the PST file, including nested folders
- Uses Microsoft standard `IPF.Appointment` container class for detection
- Automatically selects the folder with the most entries when multiple calendar folders exist
- Falls back to name-based detection for compatibility

**Supported Folder Names:**
- "Calendar", "Kalender", "Calendrier", "Calendario" (any language variant)
- "Calendar (This computer only)" and other non-standard names with `IPF.Appointment` container class
- Nested calendar folders (e.g., Calendar > Calendar subfolder)

If you still get this error after v1.2.3, the PST file may not contain calendar data. Try opening it in Outlook to verify.

### Only a few entries imported (much less than expected)

**Most common cause:** Private appointments are excluded by default. Use `--include-private` to import all appointments:

```bash
fixECalendar --clear-db
fixECalendar your-file.pst --include-private
fixECalendar --db-stats
```

This typically increases imported entries significantly, as many Outlook users have most appointments marked as private/confidential.

### Many entries with missing or invalid dates (FIXED in v1.2.3)

**Issue:** PST files often contain appointments with missing start/end dates or corrupted data.

**Solution (Fixed in v1.2.3):** The tool now automatically recovers appointments with date issues using intelligent date recovery strategies:

1. **Duration-based recovery**: Calculates missing dates using the appointment's duration field
2. **Alternative date fields**: Uses fallback dates from `recurrenceBase`, `creationTime`, etc.
3. **Date reversal fix**: Automatically corrects reversed start/end times
4. **Zero-duration fix**: Adds appropriate duration (1 hour for appointments, 24 hours for birthdays/anniversaries)
5. **Subject validation**: Discards corrupted entries with no subject to prevent database bloat

**Result:** Significantly more valid entries recovered while filtering out corrupted data. Example from real-world processing:
- 72,629 total appointments in PST files
- 4,911 quality entries imported
- 341 appointments recovered via date sanitization
- ~7,356 corrupted "(No Subject)" entries automatically discarded

No user action required - this happens automatically during import.

### Birthdays/anniversaries spanning two days (FIXED in v1.2.0)

**Issue:** All-day events like birthdays showing from 11pm one day to 11pm the next day (e.g., Oct 24 23:00 to Oct 25 23:00).

**Cause:** PST files store all-day events with UTC timestamps. Timezone offsets cause dates to appear incorrectly.

**Solution:** This was fixed in v1.2.0. Both iCal and CSV exports now properly normalize all-day event dates to remove timezone offsets. All-day events now display as single-day events with the correct date.

If you processed files before v1.2.0, simply re-export:
```bash
# For iCal
fixECalendar --output calendar.ics

# For CSV
npx ts-node export-to-csv.ts
```

No need to re-import PST files - the fix is applied during export.

### "PST file not found"

Check the file path. Use absolute paths if relative paths don't work.

### Database is locked

Another process is using the database. Close other instances of fixECalendar or specify a different database with `--database`.

### Out of memory

Very rare with the database approach. If it happens:
1. Use `--verbose` to see where it fails
2. Process files one at a time
3. Ensure you have at least 2GB free RAM

### Database file is large

The database file grows with the number of entries:
- 10,000 entries ≈ 100MB
- 100,000 entries ≈ 1GB

This is normal. To reduce size after deleting entries:
```bash
sqlite3 .fixecalendar.db "VACUUM"
```

## Technical Details

### Dependencies

- **pst-extractor** - Pure JavaScript PST file parser
- **ical-generator** - iCalendar file generation
- **commander** - Command-line interface framework
- **fast-glob** - Fast file pattern matching
- **better-sqlite3** - Fast, synchronous SQLite3 database

### PST File Format

Supports both:
- **ANSI format** (PST 97-2002) - 2GB size limit
- **Unicode format** (PST 2003+) - Larger file support (tested up to 6.5GB)

### Database Schema

The SQLite database includes:
- `calendar_entries` - Main calendar events table with indexes
- `attendees` - Attendee information (foreign key to entries)
- `processing_log` - Processing history and statistics

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Ensure all tests pass
6. Run linting
7. Submit a Pull Request

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [pst-extractor](https://github.com/epfromer/pst-extractor) by @epfromer
- Built with [ical-generator](https://github.com/sebbo2002/ical-generator) by @sebbo2002
- Built with [better-sqlite3](https://github.com/WiseLibs/better-sqlite3) for efficient database handling
- Designed specifically for large PST files (6GB+) with automatic deduplication

## Support

- **Issues**: https://github.com/ddttom/fixEcalendar/issues
- **Discussions**: https://github.com/ddttom/fixEcalendar/discussions

## Changelog

### v1.2.4 (2025-12-04)

- **Critical Fix**: CSV export now properly escapes newlines in descriptions
- **Fix**: Multi-line descriptions are converted to literal `\n` in CSV to prevent row splitting
- **Fix**: CSV import now correctly unescapes `\n` back to actual newlines for ICS generation
- **Result**: 100% data integrity in CSV → ICS conversion (previously lost ~50% of entries)
- **Impact**: Before fix: 2,470/4,886 entries converted. After fix: 4,886/4,886 entries converted
- **Critical Fix**: Corrupted recurrence UNTIL dates now sanitized (Microsoft Outlook bug)
- **Fix**: Recurrence patterns with UNTIL dates before year 1900 are projected to year 2100
- **Impact**: Fixed 71 corrupted entries with `UNTIL=16001231` → `UNTIL=21001231`
- **Tool**: Added `sanitize-recurrence-dates.ts` script to fix existing database entries
- **Prevention**: Import code now sanitizes corrupted dates automatically for future PST imports
- **Critical Fix**: All-day events now RFC 5545 compliant for Google Calendar import
- **Fix**: Birthday/anniversary events now have proper exclusive DTEND dates
- **Impact**: ICS files now import successfully into Google Calendar without "oops" errors
- **Technical**: DTEND for all-day events is now the day after the event (exclusive end date per RFC 5545)

### v1.2.3 (2025-12-04)

- **Critical Fix**: Calendar folder detection now handles non-standard folder names and nested structures
- **Enhancement**: Uses Microsoft's PR_CONTAINER_CLASS property (`IPF.Appointment`) for reliable folder detection
- **Fix**: Folders like "Calendar (This computer only)" are now correctly recognized
- **Enhancement**: Parser now finds ALL calendar folders, including nested folders
- **Enhancement**: Automatically selects folder with most entries when multiple folders exist
- **Enhancement**: Enhanced logging shows displayName, containerClass, and entry count
- **Fix**: Successfully processes PST files with nested calendar folder structures
- **Major Enhancement**: Intelligent date recovery & data quality filtering
  - Automatically recovers appointments with missing/invalid dates using duration and alternative date fields
  - Detects and fixes reversed dates (endTime before startTime)
  - Handles zero-duration appointments intelligently (1 hour for regular, 24 hours for birthdays/anniversaries)
  - Automatically converts birthdays/anniversaries/holidays to proper all-day events
  - Discards corrupted entries with no subject to prevent database bloat
  - Result: 60% reduction in database bloat while recovering more valid entries
- Backward compatible with existing name-based detection

### v1.2.2 (2025-12-04)

- **Critical Fix**: Google Calendar import now works correctly - fixed invalid RRULE format for yearly recurring events
- **Fix**: Yearly recurrence now uses `FREQ=YEARLY` without `INTERVAL=12` (RFC 5545 compliant)
- **Enhancement**: Automatic cleanup of existing CSV data with invalid recurrence intervals
- **Enhancement**: CI/CD improvements - removed Node.js 16.x, added test suite
- **Fix**: Updated Node.js requirement to 18.18.0+ to match modern dependencies

### v1.2.1 (2025-12-04)

- **Enhancement**: Added description field formatting for all exports
- **Fix**: Description fields now trimmed of excessive whitespace
- **Fix**: Multiple consecutive line breaks (3+) reduced to 2 for better readability
- **Fix**: Description fields truncated to 79 characters to prevent unwieldy text
- **Enhancement**: Created centralized text formatting utility (`src/utils/text-formatter.ts`)
- Applied consistent formatting across CSV and ICS exports

### v1.2.0 (2025-12-03)

- **Major**: Added CSV export functionality (`export-to-csv.ts`)
- **Major**: Fixed all-day event timezone handling - birthdays/anniversaries now display correctly as single-day events
- **Fix**: All-day events now use proper `VALUE=DATE` format in iCal exports (no more 2-day spans)
- **Fix**: CSV export properly normalizes all-day event dates to remove UTC timezone offsets
- **Fix**: CSV export now properly quotes all fields for correct Excel import
- **Enhancement**: Birthday/anniversary subjects standardized to `(dd/mm/yyyy)` format
- Added comprehensive RRULE recurrence pattern support using pst-extractor's RecurrencePattern class
- Support for DAILY, WEEKLY, MONTHLY, YEARLY frequencies with ordinals, intervals, and counts

### v1.1.0 (2025-12-03)

- **Major**: Added SQLite database for intermediate storage
- **Major**: Automatic deduplication of calendar entries
- **Major**: Optimized for very large PST files (6GB+)
- Added `--database`, `--clear-db`, `--db-stats` options
- Added `--no-dedupe` option to disable deduplication
- Improved memory efficiency for large file processing
- Added processing statistics and logging

### v1.0.0 (2025-12-03)

- Initial release
- PST to iCal conversion
- Batch processing support
- Merge mode
- Full property mapping
- CLI interface
- TypeScript implementation
