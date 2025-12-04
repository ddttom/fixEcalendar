# fixECalendar

Convert calendar entries from Microsoft Outlook PST files to iCalendar (iCal/ICS) format.

**Built for large PST files (6GB+) with automatic deduplication using an intermediate SQLite database.**

## Features

- ✅ Extract calendar appointments from PST files
- ✅ **Handle very large PST files (6GB+) efficiently**
- ✅ **Automatic deduplication** - never import the same event twice
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

## Installation

### Prerequisites

- Node.js 16 or higher
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
- Multi-line descriptions and special characters are handled correctly
- Commas within field data don't break the CSV structure

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
- Generates `calendar-export.ics` with complete iCal structure
- Preserves all properties (attendees, recurrence, reminders, etc.)
- Handles all-day events and birthdays correctly

**Use case:** This is useful when you want to:
- Share calendar data in iCal format after reviewing/editing the CSV
- Import into calendar applications (Apple Calendar, Google Calendar, Outlook)
- Create a portable calendar file from CSV data
- Convert between formats for different use cases

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
- `FREQ=YEARLY;INTERVAL=12;BYMONTHDAY=29` - Yearly on 29th (birthdays)

## Limitations

- **Attachments**: Not included in iCal export
- **Categories**: Not currently mapped
- **Private Appointments**: Excluded by default - **IMPORTANT:** Use `--include-private` to include personal/private/confidential appointments. Many users have most of their appointments marked as private, so this flag is highly recommended for personal calendar exports.

## Troubleshooting

### "No calendar folder found in PST file"

The PST file doesn't contain a calendar folder. The tool searches for folders named "Calendar", "Kalender", "Calendrier", or "Calendario".

### Only a few entries imported (much less than expected)

**Most common cause:** Private appointments are excluded by default. Use `--include-private` to import all appointments:

```bash
fixECalendar --clear-db
fixECalendar your-file.pst --include-private
fixECalendar --db-stats
```

This typically increases imported entries significantly, as many Outlook users have most appointments marked as private/confidential.

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
