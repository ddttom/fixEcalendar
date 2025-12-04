# CLAUDE.md

This file provides guidance to AI assistants when working with code in this repository.

## Project Overview

fixEcalendar converts Microsoft Outlook PST files to iCalendar (ICS) and CSV formats. It's designed specifically for large PST files (6GB+) with automatic deduplication using an intermediate SQLite database.

## Build and Development Commands

### Building
```bash
npm run build          # Compile TypeScript to dist/
npm run prepare        # Runs build automatically (pre-publish hook)
```

### Development
```bash
npm run dev            # Run with ts-node (no build required)
node dist/index-with-db.js [args]  # Run compiled version (production mode)
npx ts-node export-to-csv.ts       # Export database to CSV
```

### Testing
```bash
npm test               # Run all Jest tests
npm run test:watch     # Watch mode for development
npm run test:coverage  # Generate coverage report
```

### Code Quality
```bash
npm run lint           # Check code with ESLint
npm run lint:fix       # Auto-fix ESLint issues
npm run format         # Format code with Prettier
```

## Architecture Overview

### Three-Layer Pipeline

```
PST File → Parser Layer → Database Layer → Export Layer → ICS/CSV
```

1. **Parser Layer** (`src/parser/`)
   - `pst-parser.ts`: Opens PST files, finds calendar folders
   - `calendar-extractor.ts`: Extracts appointments, applies filters
   - Converts Outlook format → CalendarEntry objects

2. **Database Layer** (`src/database/`)
   - `calendar-db.ts`: SQLite persistence with deduplication
   - SHA256 hash of 6 fields (uid, subject, start/end time, location, organizer)
   - Prevents duplicate imports across multiple PST files

3. **Export Layer** (`src/converter/`)
   - `property-mapper.ts`: Maps Outlook → iCal properties
   - `ical-converter.ts`: Generates ICS files
   - `export-to-csv.ts`: Generates CSV files

### Dual CLI Entry Points

- **`src/index-with-db.ts`** (Primary): Production mode with database and deduplication
- **`src/index.ts`** (Legacy): Direct PST → ICS without database

The database mode is recommended for all use cases, especially large files and incremental processing.

### Data Flow Example

```typescript
// Processing a PST file
PSTParser.open(file)
  → CalendarExtractor.extractFromFolder()
  → CalendarEntry[] (normalized in-memory)
  → CalendarDatabase.addEntry() (with deduplication)
  → SQLite persistence

// Exporting
CalendarDatabase.getEntriesInRange()
  → PropertyMapper.mapToICalEvent()
  → ICalConverter.convert()
  → .ics or .csv file
```

## Key Design Patterns

### Deduplication System
Located in `src/database/calendar-db.ts`:
- Generates SHA256 hash from 6 fields per entry
- O(1) duplicate lookup via indexed hash column
- Enables incremental PST processing without duplicates
- Tracks source file for statistics

### Property Mapping
Located in `src/converter/property-mapper.ts`:
- Centralizes all Outlook → iCal conversions
- Special handling for all-day events (UTC midnight normalization)
- Birthday/anniversary subject standardization to `(dd/mm/yyyy)` format
- **Historical date modernization**: Recurring birthdays/anniversaries before 1970 use modern anchor date (2020) for Google Calendar compatibility
  - Original birth year preserved in subject line (e.g., "John's Birthday (15/03/1965)")
  - Only affects recurring birthdays/anniversaries, NOT actual event dates
  - Regular events keep their original dates (2000-2004 appointments preserved)
- Recurrence pattern parsing from binary RecurrencePattern class

### Text Formatting
Located in `src/utils/text-formatter.ts`:
- `formatDescription()`: Trims whitespace, normalizes line breaks, truncates to 79 chars
- **HTML/CSS junk filtering**: Detects and removes malformed Outlook extraction artifacts
- Filters patterns: "false\nfalse\nfalse", "EN-GB", "X-NONE", "/* Style Definitions */", "table.MsoNormal"
- Applied consistently across all export formats
- Prevents unwieldy text and corrupted descriptions in exports

## Database Schema

The SQLite database (`.fixecalendar.db`) has three tables:

1. **calendar_entries**: Main table with all appointment data
   - Indexed on: hash (UNIQUE), uid, start_time, source_file, (subject, start_time, end_time)
   - `hash` field enables O(1) duplicate detection

2. **attendees**: Normalized attendee data (foreign key to calendar_entries)

3. **processing_log**: Audit trail of PST file processing runs

## Critical Files for Architecture Changes

- **`src/database/calendar-db.ts`**: Modify for schema changes, add indexes, change deduplication logic
- **`src/converter/property-mapper.ts`**: Add new Outlook → iCal property mappings
- **`src/parser/calendar-extractor.ts`**: Change extraction filters, add new appointment fields
- **`src/utils/text-formatter.ts`**: Modify text cleaning/formatting rules
- **`src/parser/types.ts`**: Add fields to CalendarEntry interface (affects entire pipeline)

## Export Scripts

### CSV Export (`export-to-csv.ts`)
Standalone script that exports calendar entries from the SQLite database to CSV format.
- Reads directly from `.fixecalendar.db`
- Outputs `calendar-export.csv` with 15 columns
- Special handling for birthdays/anniversaries and all-day events
- Proper CSV escaping for Excel compatibility
- **Newline escaping**: Converts actual newlines to literal `\n` to ensure each entry is a single CSV row

### ICS Export from CSV (`export-to-ical.ts`)
Standalone script that converts CSV export to iCalendar (ICS) format.
- Reads from `calendar-export.csv`
- Parses CSV with proper quoted field handling
- **Newline unescaping**: Converts literal `\n` back to actual newlines in descriptions
- Converts to CalendarEntry objects
- Uses ICalConverter to generate RFC 5545 compliant ICS files
- Preserves all properties including attendees, recurrence, reminders
- Reuses existing PropertyMapper for consistent formatting
- **Includes automatic cleanup**: Removes invalid `INTERVAL=12` from yearly recurrence rules for Google Calendar compatibility
- **Automatic file splitting**: Splits large calendars into 499-event chunks for Google Calendar compatibility
  - **Working value**: 499 events per file tested and confirmed to import successfully
  - **Small calendars (≤499 events)**: Generates single file `calendar-export.ics`
  - **Large calendars (500+ events)**: Generates split files `calendar-part-X-of-Y.ics` with 499 events each
  - **Import instructions**: Script provides clear sequential import instructions for split files

**Workflow:**
```
Database → export-to-csv.ts → calendar-export.csv → export-to-ical.ts → ICS files (split if >499 events)
```

This two-step process allows for CSV review/editing before ICS generation.

### Recurrence Date Sanitization (`sanitize-recurrence-dates.ts`)
Utility script that fixes corrupted UNTIL dates in existing database entries.
- Scans all calendar entries with recurrence patterns containing UNTIL clauses
- Identifies corrupted dates with year < 1900 (Microsoft Outlook bug)
- Projects corrupted dates to year 2100 (same month/day)
- Reports number of entries fixed
- **Usage**: `npx ts-node sanitize-recurrence-dates.ts`
- **When needed**: After importing PST files, or if exports show `UNTIL=16001231` dates
- **Note**: v1.2.4+ automatically sanitizes during import, so this is for legacy data

**Common corrupted date pattern:**
- `UNTIL=16001231` (year 1600) → `UNTIL=21001231` (year 2100)

**Database operations:**
```typescript
// Find all entries with UNTIL dates
SELECT id, subject, recurrence_pattern
FROM calendar_entries
WHERE recurrence_pattern LIKE '%UNTIL=%'

// Update corrupted dates
UPDATE calendar_entries
SET recurrence_pattern = [sanitized_pattern]
WHERE id = [entry_id]
```

## Common Development Patterns

### Adding a New Export Format

1. Create formatter in `src/converter/` (follow `property-mapper.ts` pattern)
2. Use `CalendarDatabase.getAllEntries()` or `getEntriesInRange()` for data
3. Apply `formatDescription()` from `text-formatter.ts` for consistency
4. Add CLI option in `index-with-db.ts`

### Adding a New Outlook Property

1. Update `CalendarEntry` interface in `src/parser/types.ts`
2. Extract in `calendar-extractor.ts` → `parseAppointment()`
3. Add database column in `calendar-db.ts` schema
4. Map to iCal in `property-mapper.ts` → `mapToICalEvent()`
5. Consider if it should be part of deduplication hash

### Modifying Deduplication Logic

If changing which fields determine uniqueness:
1. Update `generateEntryHash()` in `calendar-db.ts`
2. Consider existing databases may need clearing (`--clear-db`)
3. Document breaking change in CHANGELOG.md

## Testing Notes

- PST files are binary format - use test fixtures in `tests/` directory
- Database tests should use in-memory SQLite (`:memory:`)
- Mock `pst-extractor` library for unit tests (it's a native addon)
- Integration tests require actual PST files (not in git)

## TypeScript Configuration

- Target: ES2022 (modern Node.js 18.18.0+)
- Strict mode enabled (full type safety)
- CommonJS modules (CLI compatibility)
- Source maps generated for debugging

## Node.js Version Requirements

- **Minimum Version**: Node.js 18.18.0 or higher
- **Reason**: Modern ESLint 9 and TypeScript ESLint 8 dependencies require Node 18.18.0+
- **CI/CD**: Tests run on Node.js 18.x and 20.x (16.x removed in v1.2.2)
- **Package.json**: `engines.node` field set to `>=18.18.0`

## Performance Characteristics

- **Throughput**: ~1,000-2,000 entries/second (CPU dependent)
- **Memory**: 500MB-1GB (SQLite overflows to disk)
- **Database size**: ~100MB per 10,000 entries
- **Typical 6GB PST**: 30-60 minutes processing time

## Important Implementation Details

### All-Day Event Handling
PST files store all-day events with UTC timestamps, causing timezone offset issues. The converter normalizes these to UTC midnight using `Date.UTC()` to ensure single-day display in exports.

### Birthday Subject Standardization
The system extracts the date from the recurrence pattern's `BYMONTHDAY` and formats subjects consistently as `Name (dd/mm/yyyy)`, removing various legacy date formats.

### Recurrence Patterns (Critical: Google Calendar Compatibility)
Uses `pst-extractor`'s RecurrencePattern class for binary parsing, then converts to RFC5545 RRULE format. Supports DAILY, WEEKLY, MONTHLY, YEARLY frequencies with ordinals, intervals, and counts.

**Important Fix (v1.2.2)**: Outlook stores yearly recurring events with `period=12` (12 months), but iCalendar RFC 5545 expects yearly events to use `INTERVAL=1` or omit the interval entirely. The code now automatically:
1. Detects yearly events with `period=12` in `calendar-extractor.ts`
2. Skips adding `INTERVAL=12` to the RRULE
3. Cleans up existing CSV data with invalid intervals in `export-to-ical.ts`

This ensures Google Calendar and other RFC 5545 compliant applications can import the ICS files without silent failures.

### HTML Description Processing
Descriptions are extracted from `body` or `bodyHTML` fields. HTML tags are stripped via regex, HTML entities decoded, then passed through `formatDescription()` utility.

## CLI Binary

The npm package installs a global `fixECalendar` command that points to `dist/index-with-db.js`. When making changes to the CLI:
1. Modify `src/index-with-db.ts`
2. Run `npm run build`
3. Test with `node dist/index-with-db.js [args]`

## Known Issues and Important Fixes

### CSV Export Newline Handling (Fixed in v1.2.4)
**Issue**: CSV export with multi-line descriptions created broken CSV files where entries spanned multiple rows, causing the CSV → ICS conversion to lose ~50% of entries.

**Root Cause**: The `escapeCSV()` function only escaped quotes but didn't escape newlines. When descriptions contained `\n` characters, they created actual line breaks in the CSV file. The CSV import script used a simple `split('\n')` which treated each line as a separate row, breaking entries that should span a single row into multiple invalid rows.

**Solution (Fixed in v1.2.4)**:
1. **CSV Export** (`export-to-csv.ts:21`): Added `.replace(/\r?\n/g, '\\n')` to convert actual newlines to literal `\n` text
2. **CSV Import** (`export-to-ical.ts:60`): Added `.replace(/\\n/g, '\n')` to convert literal `\n` back to actual newlines

**Impact**:
- **Before fix**: CSV had 26,925 rows (broken), only 2,470 entries imported to ICS (50% data loss)
- **After fix**: CSV has 4,886 rows (correct), all 4,886 entries imported to ICS (100% data integrity)

**Code Pattern**:
```typescript
// In export-to-csv.ts
function escapeCSV(value: any): string {
  const str = String(value);
  // Escape quotes by doubling them and replace newlines with literal \n
  return `"${str.replace(/"/g, '""').replace(/\r?\n/g, '\\n')}"`;
}

// In export-to-ical.ts
const description = fields[6]?.replace(/\\n/g, '\n');
```

**Location**:
- `export-to-csv.ts:12-22` (escapeCSV function)
- `export-to-ical.ts:60` (description unescaping)

### Corrupted Recurrence UNTIL Dates (Fixed in v1.2.4)
**Issue**: Microsoft Outlook PST files sometimes contain recurrence patterns with corrupted UNTIL dates set to year 1600 (e.g., `UNTIL=16001231`). This causes calendar applications to either reject the recurrence rule or display events incorrectly.

**Root Cause**: Known Microsoft Outlook bug where recurrence end dates are stored with invalid year values (< 1900). Affected approximately 71 entries in typical PST files.

**Solution (Fixed in v1.2.4)**:
1. **Import sanitization** (`src/parser/calendar-extractor.ts:440-455`): Projects corrupted dates (year < 1900) to year 2100 during PST import
2. **Database cleanup tool** (`sanitize-recurrence-dates.ts`): One-time script to fix existing database entries

**Impact**:
- **Before fix**: 71 entries with `UNTIL=16001231` (year 1600 - invalid)
- **After fix**: All sanitized to `UNTIL=21001231` (year 2100 - reasonable end date)

**Code Pattern**:
```typescript
// In calendar-extractor.ts
case 8225: // AfterDate
  if (pattern.endDate) {
    let endYear = pattern.endDate.getFullYear();
    // Sanitize corrupted end dates: project dates before 1900 to year 2100
    if (endYear < 1900) {
      endYear = 2100;
      logger.info(`Sanitized corrupted recurrence end date from ${pattern.endDate.getFullYear()} to 2100`);
    }
    // Cap end date at 2100 for consistency
    if (endYear > 2100) {
      endYear = 2100;
    }
    const month = String(pattern.endDate.getMonth() + 1).padStart(2, '0');
    const day = String(pattern.endDate.getDate()).padStart(2, '0');
    rruleParts.push(`UNTIL=${endYear}${month}${day}`);
  }
  break;
```

**Location**:
- `src/parser/calendar-extractor.ts:438-456` (sanitization during import)
- `sanitize-recurrence-dates.ts` (database cleanup script)

**Usage**: Run `npx ts-node sanitize-recurrence-dates.ts` to fix existing database entries, then re-export CSV/ICS files.

### Absurd Recurrence Pattern Validation (Fixed in v1.2.5)
**Issue**: Events with corrupted recurrence patterns were creating absurd recurring schedules. For example, "Toms Online Photo Course" (a 4-day event) had recurrence pattern `FREQ=DAILY;UNTIL=21001231T235959Z`, making it repeat **every single day for 76 years** (27,740 occurrences). This completely destroyed calendar views by filling them with thousands of duplicate entries.

**Root Cause**: The sanitization logic for corrupted UNTIL dates (year < 1900) projected them to year 2100 without validating if the resulting span was reasonable. Microsoft Outlook PST files contain corrupted recurrence end dates that are meaningless, but the code treated them as legitimate and created 70-85 year recurring patterns. **72 entries** were affected with UNTIL=2100 spanning unreasonable time periods.

**Solution (Fixed in v1.2.5)**:
1. **Recurrence Validator Module** ([src/utils/recurrence-validator.ts](src/utils/recurrence-validator.ts)): Core validation logic with configurable caps per frequency
2. **PST Import Validation** ([src/parser/calendar-extractor.ts:259-266, 471-503](src/parser/calendar-extractor.ts#L259-L266)): Validates patterns during extraction and caps sanitized UNTIL dates
3. **COUNT Preference Logic** ([src/parser/calendar-extractor.ts:471-484](src/parser/calendar-extractor.ts#L471-L484)): Prefers COUNT over corrupted UNTIL when occurrenceCount ≤ 50
4. **Database Cleanup Script** ([cleanup-suspicious-recurrence.ts](cleanup-suspicious-recurrence.ts)): One-time script to fix existing database entries
5. **CSV Import Cleanup** ([export-to-ical.ts:107-141](export-to-ical.ts#L107-L141)): Defense-in-depth cleanup during CSV→ICS conversion

**Validation Rules**:
- **DAILY**: Maximum 5 years (unless infinite/COUNT-based)
- **WEEKLY**: Maximum 10 years
- **MONTHLY**: Maximum 20 years
- **YEARLY**: Maximum 100 years (birthdays/anniversaries)
- **Single occurrence**: Strip recurrence if `occurrenceCount === 1`
- **Prefer COUNT over UNTIL**: When UNTIL is corrupted (year < 1900) AND occurrenceCount exists (1-50), use COUNT for accuracy

**Impact**:
- **Before fix**: "Toms Online Photo Course" repeated daily for 76 years (27,740 occurrences)
- **After fix**: Changed to `COUNT=4` (4 occurrences - the actual 4-day course duration)
- **Affected**: 72 entries with UNTIL=2100 (11 daily, 61 weekly/monthly/yearly)
- **COUNT-based**: 202 entries now use precise COUNT instead of capped UNTIL dates
- **Result**: Clean calendar exports without absurd recurring patterns and accurate occurrence counts

**Code Pattern**:
```typescript
// In calendar-extractor.ts (COUNT preference logic - lines 471-484)
// Special case: If endDate is corrupted (year < 1900) but occurrenceCount is small and reasonable,
// prefer COUNT over UNTIL. This handles cases like "4-day course" stored with corrupted end date.
if (
  endYear < 1900 &&
  pattern.occurrenceCount &&
  pattern.occurrenceCount > 1 &&
  pattern.occurrenceCount <= 50
) {
  logger.info(
    `Using COUNT=${pattern.occurrenceCount} instead of corrupted UNTIL date (year ${originalYear}) for better accuracy`
  );
  rruleParts.push(`COUNT=${pattern.occurrenceCount}`);
  break; // Skip UNTIL processing
}

// In calendar-extractor.ts (PST import validation)
const rawRRule = this.buildRRuleFromPattern(pattern, startTime);

// Validate recurrence pattern to catch corrupted/unreasonable patterns
const validatedRRule = RecurrenceValidator.validateRecurrencePattern(
  pattern,
  startTime,
  appointment.subject || 'Untitled',
  rawRRule
);

return validatedRRule || undefined;

// In calendar-extractor.ts (sanitized UNTIL date validation - lines 485-503)
if (endYear < 1900) {
  const originalYear = endYear;
  endYear = 2100;

  // Validate the sanitized date is reasonable
  const yearsSpan = endYear - startTime.getFullYear();
  const maxYears = RecurrenceValidator.getMaxYearsForFrequency(pattern.recurFrequency);

  if (yearsSpan > maxYears) {
    endYear = startTime.getFullYear() + maxYears;
    logger.warn(`Capping sanitized recurrence to ${endYear} (${maxYears} years max for this frequency)`);
  }
}

// In export-to-ical.ts (CSV import cleanup)
if (recurrencePattern.includes('UNTIL=2100')) {
  const yearsSpan = 2100 - startDateObj.getFullYear();

  if (recurrencePattern.includes('FREQ=DAILY') && yearsSpan > 5) {
    const cappedYear = startDateObj.getFullYear() + 5;
    recurrencePattern = recurrencePattern.replace(/UNTIL=2100\d{4}T\d{6}Z/, `UNTIL=${cappedYear}...`);
  }
}
```

**Location**:
- [src/utils/recurrence-validator.ts](src/utils/recurrence-validator.ts) - Core validation logic (NEW)
- [src/parser/calendar-extractor.ts:259-266](src/parser/calendar-extractor.ts#L259-L266) - PST import validation
- [src/parser/calendar-extractor.ts:466-502](src/parser/calendar-extractor.ts#L466-L502) - Sanitized date validation
- [cleanup-suspicious-recurrence.ts](cleanup-suspicious-recurrence.ts) - Database cleanup script (NEW)
- [export-to-ical.ts:107-141](export-to-ical.ts#L107-L141) - CSV import cleanup
- [src/config/constants.ts:34-49](src/config/constants.ts#L34-L49) - Configuration constants

**Usage**:
```bash
# Fix existing database entries
npx ts-node cleanup-suspicious-recurrence.ts

# Review the cleanup report
cat recurrence-cleanup-report.csv

# Re-export with fixed patterns
npx ts-node export-to-csv.ts
npx ts-node export-to-ical.ts
```

**Configuration**: Validation thresholds can be customized in [src/config/constants.ts](src/config/constants.ts) under `RECURRENCE_VALIDATION_CONFIG`.

### RFC 5545 All-Day Event Compliance (Fixed in v1.2.4)
**Issue**: Google Calendar rejected ICS imports with "oops we could not import this file" error due to invalid all-day event formatting. Birthday and anniversary events had identical DTSTART and DTEND dates (e.g., `DTSTART=19260613, DTEND=19260613`), violating RFC 5545 which requires DTEND to be exclusive (the day after the event ends).

**Root Cause**: The CSV export stores all-day events with the same start and end date (e.g., June 13 to June 13), which is intuitive from a user perspective. However, RFC 5545 requires DTEND to be the day AFTER the event ends (exclusive). Birthday/anniversary events were not adjusting the end date during ICS conversion, while regular all-day events already had the correct exclusive end date from the database.

**Solution (Fixed in v1.2.4)**:
Modified `src/converter/property-mapper.ts:74` to add +1 day to the end date for birthday/anniversary events only. Regular all-day events already have correct exclusive dates from the CSV export.

**Impact**:
- **Before fix**: `DTSTART;VALUE=DATE:19260613` / `DTEND;VALUE=DATE:19260613` ❌ (violates RFC 5545)
- **After fix**: `DTSTART;VALUE=DATE:19260613` / `DTEND;VALUE=DATE:19260614` ✓ (RFC 5545 compliant)
- **Result**: ICS files now import successfully into Google Calendar, Apple Calendar, and all RFC 5545 compliant applications

**Code Pattern**:
```typescript
// In property-mapper.ts (birthday/anniversary handling)
if (isBirthdayOrAnniversary && entry.recurrencePattern) {
  const byMonthDayMatch = entry.recurrencePattern.match(/BYMONTHDAY=(\d+)/);
  if (byMonthDayMatch) {
    const correctDay = parseInt(byMonthDayMatch[1]);
    startTime = new Date(year, month, correctDay);
    endTime = new Date(year, month, correctDay + 1); // RFC 5545: DTEND is exclusive
    isAllDay = true;
  }
}

// Regular all-day events (no +1 needed - CSV already has exclusive dates)
if (isAllDay) {
  startTime = new Date(Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()));
  endTime = new Date(Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()));
}
```

**Location**: `src/converter/property-mapper.ts:63-94`

**Key Insight**: The CSV export treats birthday/anniversary events differently from regular all-day events:
- **Birthdays/anniversaries**: CSV has same start and end date (June 13 → June 13) → needs +1 in ICS
- **Regular all-day events**: CSV already has exclusive end date (May 4 → May 5) → no adjustment needed

### RFC 5545 Recurring Event Compliance (Fixed in v1.2.4)
**Issue**: Recurring events with UNTIL clauses were not importing into Google Calendar because:
1. UNTIL values lacked time components (e.g., `UNTIL=21001231` instead of `UNTIL=21001231T235959Z`)
2. Daily recurrence intervals were stored in minutes instead of days (e.g., `INTERVAL=1440` instead of `INTERVAL=1`)

**Root Cause**:
1. **UNTIL format**: RFC 5545 requires that UNTIL must match the format of DTSTART. Since DTSTART includes a time component (`DTSTART:20220915T120000Z`), UNTIL must also include a time component in UTC format
2. **Daily intervals**: Microsoft Outlook PST files store daily recurrence periods as minutes (per MS-OXOCAL specification), where 1440 minutes = 1 day. The code was directly using these minute values as day intervals, causing `INTERVAL=1440` to mean "every 1440 days" instead of "every 1 day"

**Solution (Fixed in v1.2.4)**:
1. **UNTIL time component** (`calendar-extractor.ts:455`, `export-to-ical.ts:80-82`):
   - Added `T235959Z` suffix to all UNTIL values during import
   - Added regex cleanup in CSV import to convert date-only UNTIL to date-time format
2. **Daily interval conversion** (`calendar-extractor.ts:380-384`, `export-to-ical.ts:84-93`):
   - Convert Outlook's minute-based intervals to days by dividing by 1440
   - Clean up existing CSV data during ICS conversion

**Impact**:
- **Before fix**: 72 recurring events with `UNTIL=21001231` (rejected by Google Calendar)
- **After fix**: All UNTIL values have `UNTIL=21001231T235959Z` format ✓
- **Before fix**: `INTERVAL=1440` meant every 1440 days (~4 years)
- **After fix**: `INTERVAL=1` means every 1 day ✓
- **Result**: All recurring events including Genealogy entries now import successfully

**Code Pattern**:
```typescript
// In calendar-extractor.ts (UNTIL with time)
const month = String(pattern.endDate.getMonth() + 1).padStart(2, '0');
const day = String(pattern.endDate.getDate()).padStart(2, '0');
rruleParts.push(`UNTIL=${endYear}${month}${day}T235959Z`);

// In calendar-extractor.ts (daily interval conversion)
if (pattern.recurFrequency === 8202) { // Daily
  const days = Math.floor(pattern.period / 1440);
  if (days > 1) {
    rruleParts.push(`INTERVAL=${days}`);
  }
}

// In export-to-ical.ts (UNTIL cleanup)
if (recurrencePattern?.includes('UNTIL=')) {
  recurrencePattern = recurrencePattern.replace(/UNTIL=(\d{8})(?!T)/g, 'UNTIL=$1T235959Z');
}

// In export-to-ical.ts (interval cleanup)
if (recurrencePattern?.includes('FREQ=DAILY') && recurrencePattern?.includes('INTERVAL=')) {
  recurrencePattern = recurrencePattern.replace(/INTERVAL=(\d+)/g, (match, minutes) => {
    const days = Math.floor(parseInt(minutes) / 1440);
    return days > 1 ? `INTERVAL=${days}` : '';
  });
}
```

**Location**:
- `src/parser/calendar-extractor.ts:451-455` (UNTIL format)
- `src/parser/calendar-extractor.ts:376-394` (daily interval conversion)
- `export-to-ical.ts:77-93` (CSV cleanup for both issues)

**References**:
- [RFC 5545 §3.3.10 - UNTIL must match DTSTART format](https://icalendar.org/iCalendar-RFC-5545/3-3-10-recurrence-rule.html)
- [MS-OXOCAL §2.2.1.44.1 - Daily period stored in minutes](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocal/cf7153b4-f8b5-4cb6-bf14-e78d21f94814)

### RFC 5545 BYMONTH Parameter for Yearly Events (Fixed in v1.2.4)
**Issue**: Google Calendar rejected ICS imports with "oops we could not import this file" error for yearly recurring events because they lacked the BYMONTH parameter. Birthday/anniversary events had format `FREQ=YEARLY;BYMONTHDAY=13` without specifying which month, creating RFC 5545 ambiguity.

**Root Cause**: While RFC 5545 technically allows BYMONTHDAY without BYMONTH for yearly frequency (defaulting to all months), Google Calendar and other strict parsers reject this as ambiguous. The calendar-extractor was only adding BYMONTHDAY from the pattern but not extracting the month from the event's start date.

**Solution (Fixed in v1.2.4)**:
1. **Import enhancement** (`calendar-extractor.ts:355-425`): Modified `buildRRuleFromPattern()` to accept startTime parameter and extract month from start date
2. **CSV cleanup** (`export-to-ical.ts:95-105`): Added cleanup for legacy CSV data to insert BYMONTH from start date

**Impact**:
- **Before fix**: `FREQ=YEARLY;BYMONTHDAY=13` (ambiguous - rejected by Google Calendar)
- **After fix**: `FREQ=YEARLY;BYMONTH=6;BYMONTHDAY=13` (explicit - imports successfully)
- **Result**: All 465 yearly recurring events now import successfully into Google Calendar

**Code Pattern**:
```typescript
// In calendar-extractor.ts (lines 418-425)
if (typeof pattern.patternTypeSpecific === 'number') {
  // For yearly recurrence with BYMONTHDAY, RFC 5545 requires BYMONTH for clarity
  // Extract month from startTime (1-12 where 1=January, 12=December)
  if (isYearlyFrequency) {
    const month = startTime.getMonth() + 1; // getMonth() returns 0-11, we need 1-12
    rruleParts.push(`BYMONTH=${month}`);
  }
  rruleParts.push(`BYMONTHDAY=${pattern.patternTypeSpecific}`);
}

// In export-to-ical.ts (lines 95-104)
if (recurrencePattern?.includes('FREQ=YEARLY') && recurrencePattern?.includes('BYMONTHDAY=')) {
  if (!recurrencePattern.includes('BYMONTH=')) {
    const month = parseInt(startDate.substring(5, 7)); // Extract MM from YYYY-MM-DD
    recurrencePattern = recurrencePattern.replace(/BYMONTHDAY=/, `BYMONTH=${month};BYMONTHDAY=`);
  }
}
```

**Location**:
- `src/parser/calendar-extractor.ts:355-425` (buildRRuleFromPattern with startTime parameter)
- `export-to-ical.ts:95-105` (CSV cleanup for legacy data)

**References**:
- [RFC 5545 §3.3.10 - Recurrence Rule](https://icalendar.org/iCalendar-RFC-5545/3-3-10-recurrence-rule.html)
- [iCalendar YEARLY rrule without BYMONTH - Stack Overflow](https://stackoverflow.com/questions/43084211/icalendar-yearly-rrule-without-bymonth-value)

### Description Field HTML/CSS Junk Filtering (Fixed in v1.2.4)
**Issue**: 1,832 calendar entries had corrupted descriptions containing HTML/CSS artifacts from malformed Outlook body extraction, causing Google Calendar to reject the ICS file or display garbled content.

**Root Cause**: The `extractDescription()` method in calendar-extractor.ts strips HTML tags but doesn't catch certain Microsoft Word/Outlook HTML artifacts like CSS comments, language codes, and boolean strings. These patterns appear when bodyHTML field extraction fails partially.

**Common patterns**:
- `false\nfalse\nfalse\n\nEN-GB\nX-NONE\nX-NONE\n\n/* Style Definitions */\ntable.MsoNormal`
- Language codes: EN-GB, X-NONE
- CSS comments: `/* Style Definitions */`
- MS Word artifacts: `table.MsoNormal`
- Lone boolean strings: true, false

**Solution (Fixed in v1.2.4)**:
Enhanced `formatDescription()` in `text-formatter.ts:16-64` to detect and filter out these patterns before export.

**Impact**:
- **Before fix**: 1,832 entries with corrupted descriptions
- **After fix**: All junk descriptions converted to empty strings
- **Result**: Clean ICS files that import successfully into all calendar applications

**Code Pattern**:
```typescript
// In text-formatter.ts (lines 21-47)
export function formatDescription(description: string | undefined): string {
  if (!description) return '';

  const junkPatterns = [
    /^(false\s*)+$/i,                    // Multiple "false" entries
    /^(EN-GB|X-NONE)\s*$/i,              // Language codes
    /\/\*\s*Style\s+Defin?itions?\s*\*\//i, // CSS comments
    /table\.MsoNormal/i,                 // MS Word HTML artifacts
    /^\s*(true|false)\s*$/i,             // Lone boolean strings
  ];

  const trimmed = description.trim();
  for (const pattern of junkPatterns) {
    if (pattern.test(trimmed)) return ''; // Return empty for pure junk
  }

  // Check for junk at start of description
  if (trimmed.startsWith('false\n') ||
      trimmed.startsWith('EN-GB') ||
      trimmed.startsWith('X-NONE') ||
      trimmed.includes('/* Style Defin')) {
    return '';
  }

  // Normal formatting continues...
}
```

**Location**: `src/utils/text-formatter.ts:16-64`

### Google Calendar Import (Fixed in v1.2.2)
**Issue**: ICS files generated before v1.2.2 fail to import into Google Calendar silently (no error, no events imported).

**Root Cause**: Outlook stores yearly recurring events with `RecurrencePattern.period = 12` (representing 12 months). The code was directly converting this to `INTERVAL=12` in the RRULE, but RFC 5545 specifies that yearly events should use `INTERVAL=1` or omit the interval parameter entirely.

**Solution Locations**:
1. `src/parser/calendar-extractor.ts:368` - Skips adding INTERVAL for yearly events with period=12
2. `export-to-ical.ts:71-75` - Cleans up invalid intervals when reading from CSV

**Code Pattern**:
```typescript
// In calendar-extractor.ts
if (pattern.recurFrequency === 8205 && pattern.period === 12) {
  // 12 months = 1 year, so INTERVAL=1 (which is default, no need to add)
} else {
  rruleParts.push(`INTERVAL=${pattern.period}`);
}

// In export-to-ical.ts
if (recurrencePattern?.includes('FREQ=YEARLY') && recurrencePattern?.includes('INTERVAL=12')) {
  recurrencePattern = recurrencePattern.replace(/INTERVAL=12;?/, '');
}
```

### CSV Deduplication (Fixed in v1.2.2)
Birthday and anniversary date normalization can create duplicate entries. The CSV export now tracks unique entries using a composite key (subject + start date + start time) to prevent duplicates after date standardization.

**Location**: `export-to-csv.ts:102-165`

### Split File Generation Strategy (Enhanced in v1.2.4)
**Feature**: Automatically splits large calendars into 499-event chunks for Google Calendar compatibility.

**Background**: Files with 499 events were tested and confirmed to import successfully into Google Calendar. The value of 499 is a working estimate based on successful testing, not a confirmed hard limit from Google.

**Solution**: The export-to-ical.ts script automatically handles file splitting:
1. **Small calendars (≤499 events)**: Creates single file `calendar-export.ics`
2. **Large calendars (500+ events)**: Creates split files `calendar-part-X-of-Y.ics` with 499 events each

**Implementation** (`export-to-ical.ts:142-187`):
```typescript
async function splitAndSaveCalendar(
  entries: CalendarEntry[],
  eventsPerFile: number = 499
): Promise<void> {
  const totalEvents = entries.length;
  const numFiles = Math.ceil(totalEvents / eventsPerFile);

  console.log(`\nGenerating ${numFiles} iCalendar file${numFiles > 1 ? 's' : ''} for Google Calendar...`);
  console.log(`Total events: ${totalEvents} (${eventsPerFile} events per file)\n`);

  for (let fileNum = 0; fileNum < numFiles; fileNum++) {
    const startIdx = fileNum * eventsPerFile;
    const endIdx = Math.min((fileNum + 1) * eventsPerFile, totalEvents);
    const chunkEntries = entries.slice(startIdx, endIdx);

    const outputPath = numFiles === 1
      ? 'calendar-export.ics'
      : `calendar-part-${fileNum + 1}-of-${numFiles}.ics`;

    const converter = new ICalConverter();
    const calendarName = numFiles === 1
      ? 'Exported Calendar'
      : `Exported Calendar (Part ${fileNum + 1} of ${numFiles})`;

    const calendar = converter.convert(chunkEntries, {
      calendarName,
      timezone: 'UTC',
    });

    await converter.saveToFile(calendar, outputPath);
    console.log(`✓ Created ${outputPath} with ${chunkEntries.length} events`);
  }

  console.log(`\n✓ Successfully created ${numFiles} file${numFiles > 1 ? 's' : ''}`);

  if (numFiles > 1) {
    console.log(`\n📋 Import Instructions:`);
    console.log(`   Import split files sequentially into Google Calendar:`);
    console.log(`     1. Import calendar-part-1-of-${numFiles}.ics`);
    console.log(`     2. Wait for import to complete`);
    console.log(`     3. Import calendar-part-2-of-${numFiles}.ics`);
    console.log(`     4. Repeat for all ${numFiles} files`);
  }
}
```

**User Experience**:
- **Small calendars**: Get single `calendar-export.ics` file with clear success message
- **Large calendars**: Get split files with clear import instructions
- **Sequential import**: Must import split files one at a time, waiting for each to complete

**Example Output (Large Calendar)**:
```
Generating 10 iCalendar files for Google Calendar...
Total events: 4886 (499 events per file)

✓ Created calendar-part-1-of-10.ics with 499 events
✓ Created calendar-part-2-of-10.ics with 499 events
✓ Created calendar-part-3-of-10.ics with 499 events
✓ Created calendar-part-4-of-10.ics with 499 events
✓ Created calendar-part-5-of-10.ics with 499 events
✓ Created calendar-part-6-of-10.ics with 499 events
✓ Created calendar-part-7-of-10.ics with 499 events
✓ Created calendar-part-8-of-10.ics with 499 events
✓ Created calendar-part-9-of-10.ics with 499 events
✓ Created calendar-part-10-of-10.ics with 395 events

✓ Successfully created 10 files

📋 Import Instructions:
   Import split files sequentially into Google Calendar:
     1. Import calendar-part-1-of-10.ics
     2. Wait for import to complete
     3. Import calendar-part-2-of-10.ics
     4. Repeat for all 10 files
```

**Example Output (Small Calendar)**:
```
Generating 1 iCalendar file for Google Calendar...
Total events: 250 (499 events per file)

✓ Created calendar-export.ics with 250 events

✓ Successfully created 1 file

You can import calendar-export.ics into Google Calendar or any RFC 5545-compliant calendar application.
```

**Technical Details**:
- Default chunk size: 499 events (tested and confirmed to import successfully)
- Each file is a complete, valid iCalendar with proper headers and footers
- Calendar names include part numbers for identification in the target application
- Files are generated sequentially to avoid memory issues with very large calendars
- No complete/unsplit file generated for large calendars

### Historical Birthday Date Modernization (Added in v1.2.4)
**Feature**: Automatically modernizes very old recurring birthday/anniversary dates for Google Calendar compatibility while preserving original birth years and all actual event dates.

**Background**: Google Calendar has display issues with recurring events starting before 1970. However, we cannot modernize all historical dates because many are actual appointments from 2000-2004 that must retain their accurate dates.

**Solution**: Selective date modernization that only affects recurring birthdays/anniversaries before 1970:
- **Recurring birthdays/anniversaries before 1970**: Start date changed to 2020 (modern anchor date)
- **Original birth year**: Preserved in subject line (e.g., "John's Birthday (15/03/1965)")
- **Regular events**: ALL actual event dates preserved unchanged (including 2000-2004 appointments)
- **Threshold**: Only dates before 1970 are modernized (1970-1999 dates remain unchanged)

**Implementation** (`src/converter/property-mapper.ts:59-86`):
```typescript
// Special handling for birthdays/anniversaries - force them to be all-day events with correct date
const subject = (entry.subject || '').toLowerCase();
const isBirthdayOrAnniversary =
  subject.includes('birthday') ||
  subject.includes('bithday') ||  // Handle typos
  subject.includes('anniversary');

if (isBirthdayOrAnniversary && entry.recurrencePattern) {
  // Extract day from BYMONTHDAY in recurrence pattern
  const byMonthDayMatch = entry.recurrencePattern.match(/BYMONTHDAY=(\d+)/);
  if (byMonthDayMatch) {
    const correctDay = parseInt(byMonthDayMatch[1]);
    const baseDate = new Date(entry.startTime);
    const originalYear = baseDate.getFullYear();
    const month = baseDate.getMonth();

    // For recurring birthdays/anniversaries, use a modern year as the "anchor" date
    // The original birth year is preserved in the subject line (e.g., "John's Birthday (15/03/1965)")
    // This avoids Google Calendar's issues with very old recurring event start dates
    // while maintaining accurate annual recurrence going forward
    const modernYear = originalYear < 1970 ? 2020 : originalYear;

    startTime = new Date(modernYear, month, correctDay);
    endTime = new Date(modernYear, month, correctDay + 1); // RFC 5545: DTEND is exclusive
    isAllDay = true; // Force as all-day event
  }
}
```

**Impact**:
- **Birthdays before 1970**: Modernized to 2020 anchor dates (e.g., 1926 → 2020)
- **Birthdays 1970-1999**: Keep original year (e.g., 1985 remains 1985)
- **Regular events**: All preserved with original dates (780+ events from 2000-2004)
- **Result**: All recurring birthdays import successfully while maintaining historical accuracy

**CALSCALE and METHOD Headers** (`src/converter/ical-converter.ts:26-32`):
Added required RFC 5545 headers for Google Calendar compatibility:
```typescript
const calendar = ical({
  name: options.calendarName || DEFAULT_CALENDAR_NAME,
  prodId: options.productId || `//${APP_NAME}//v${APP_VERSION}//EN`,
  timezone: options.timezone || DEFAULT_TIMEZONE,
  scale: 'GREGORIAN',
  method: ICalCalendarMethod.PUBLISH,
});
```

**Location**:
- `src/converter/property-mapper.ts:59-86` (selective date modernization)
- `src/converter/ical-converter.ts:26-32` (CALSCALE and METHOD headers)

**Key Design Decision**: The 1970 threshold was chosen based on user feedback that "we cannot have all dates 2000 or later" - preserving the distinction between very old recurring birthdays (which need modernization) and actual historical events (which must remain accurate).

### Calendar Folder Detection (Fixed in v1.2.3)
**Issue**: Folders with non-standard names like "Calendar (This computer only)" were not recognized as calendar folders, causing "No calendar folder found in PST file" errors. Additionally, nested calendar folder structures were not fully explored, and only the largest folder was processed when multiple calendar folders existed.

**Root Cause**: The parser used exact string matching on folder display names against a hardcoded list (`CALENDAR_FOLDER_NAMES`). This failed for folders with suffixes, localized names, or custom naming. The parser also stopped searching after finding the first calendar folder, missing nested folders with actual content. Initially, the fix selected only the folder with the most entries, but this missed data in smaller calendar folders like imported .ics files.

**Solution**: Implemented Microsoft's recommended approach using the PR_CONTAINER_CLASS property (containerClass in pst-extractor):
- **Primary detection**: Check if `folder.containerClass` starts with "IPF.Appointment" (the Microsoft standard for calendar folders)
- **Fallback detection**: Check display name against known calendar folder names (preserves backward compatibility)
- **Complete search**: Parser now finds ALL calendar folders in the entire PST structure, including nested folders
- **Process all folders**: Changed from selecting one folder to processing ALL calendar folders found in a PST file
- **Progress tracking**: Clear folder-by-folder progress indication when processing multiple folders
- **Nested folder support**: Handles complex folder structures like:
  ```
  📁 Calendar (This computer only) - 0 entries (processed, 0 entries found)
    └─ 📁 Calendar (This computer only) - 18,415 entries (processed, 15,987 entries extracted)
  📁 Calendar - 2 entries (processed, 2 entries extracted)
  📁 schedule.ics - 5 entries (processed, 5 entries extracted)
  ```

**Location**:
- `src/parser/pst-parser.ts:71-89` (getAllCalendarFolders method)
- `src/parser/pst-parser.ts:91-143` (findAllCalendarFolders method)
- `src/index-with-db.ts:62-130` (processWithDatabase function - processes all folders)

**Code Pattern**:
```typescript
// Find ALL calendar folders (depth-first search)
private findAllCalendarFolders(folder: any): any[] {
  const calendarFolders: any[] = [];

  // Check if current folder is a calendar folder
  if (containerClass?.trim().toLowerCase().startsWith('ipf.appointment')) {
    calendarFolders.push(folder);
  }

  // Always search subfolders (even if current folder is a calendar)
  if (folder.hasSubfolders) {
    const subFolders = folder.getSubFolders();
    for (const subFolder of subFolders) {
      calendarFolders.push(...this.findAllCalendarFolders(subFolder));
    }
  }

  return calendarFolders;
}

// Select folder with most entries
const folderWithMostEntries = folders.reduce((best, current) => {
  return (current.contentCount || 0) > (best.contentCount || 0) ? current : best;
});
```

**Enhanced Logging**: Debug output now shows displayName, containerClass, AND entry count for easier troubleshooting of folder detection issues.

### Intelligent Date Recovery & Data Quality (Added in v1.2.3)
**Feature**: Automatically recovers appointments with missing or invalid dates while filtering out corrupted entries.

**Date Recovery Strategies** (`src/parser/calendar-extractor.ts:547-636`):
1. **Duration-based recovery**: If one date exists, calculates the other using the `duration` field (in minutes)
2. **Alternative date fields**: Falls back to `recurrenceBase`, `attendeeCriticalChange`, `creationTime`, `clientSubmitTime`, or `messageDeliveryTime`
3. **Date reversal fix**: Automatically swaps dates if `endTime` is before `startTime`
4. **Zero-duration handling**:
   - **Birthdays/anniversaries/holidays**: Converted to full 24-hour all-day events (midnight to midnight)
   - **Regular appointments**: Add 1-hour duration
5. **Subject validation**: Discards entries with no subject (corrupted/incomplete data)

**Birthday/Anniversary Detection**:
Intelligently detects special events by subject keywords:
- birthday, anniversary, holiday
- easter, christmas, bank holiday, good friday
- st., saint (e.g., St. Patrick's Day)

These are automatically converted to all-day events with proper midnight-to-midnight timing.

**Data Quality Impact**:
- **Before**: 12,267 entries (including 7,356 corrupted "(No Subject)" entries)
- **After**: 4,911 quality entries (60% reduction in database bloat)
- **All-day events**: 92 birthdays/anniversaries properly formatted
- **Recovery success**: 341 appointments recovered with date sanitization

**Example Log Output**:
```
[INFO] Made "Birthday - Bill Bray" an all-day event (birthday/anniversary/holiday)
[INFO] Made "Anniversary - George & Lilian (1954)" an all-day event (birthday/anniversary/holiday)
[INFO] EndTime equals startTime for "Private Appointment", adding 1 hour duration
[INFO] Recovered endTime for "Team Meeting" using duration (60 min)
```

**Location**: `src/parser/calendar-extractor.ts:67-84` (subject validation), `src/parser/calendar-extractor.ts:547-636` (sanitizeDates method)

### File Status Report (Added in v1.2.3)
**Feature**: When processing multiple PST files, the CLI now generates a comprehensive File Status Report showing problematic files at the end of processing.

**Categories Tracked**:
1. **Files with errors**: PST files that failed to process (unreadable, corrupted, no calendar folders found) with specific error messages
2. **Files with zero entries**: Valid PST files that contained no calendar entries (e.g., empty calendars, contact-only PST files)
3. **Files with only duplicates**: PST files where all entries were duplicates of entries already in the database

**Use Case**: Helps identify which PST files need attention when batch processing multiple files. For example:
- Contact PST files mixed with calendar PST files → shown in "errors" (no calendar folder)
- Backup copies of the same PST → shown in "only duplicates"
- Empty or archived PST files → shown in "zero entries"

**Location**: `src/index-with-db.ts:270-336` (tracking and report generation)

**Example Output**:
```
=== File Status Report ===

Files with errors (3):
  - contacts.pst: No calendar folder found in PST file
  - corrupt.pst: Failed to open PST file
  - encrypted.pst: The PST file may be corrupted, encrypted, or in an unsupported format

Files with zero entries (1):
  - empty-archive.pst

Files with only duplicates (1):
  - backup-copy.pst: 15987 entries were duplicates
```

**Implementation Pattern**:
```typescript
// Track problematic files during processing
const filesWithErrors: Array<{ file: string; error: string }> = [];
const filesWithZeroEntries: string[] = [];
const filesWithOnlyDuplicates: Array<{ file: string; count: number }> = [];

for (const inputPath of inputFiles) {
  try {
    const result = await processWithDatabase(inputPath, database, options);

    if (result.found === 0) {
      filesWithZeroEntries.push(fileName);
    } else if (result.added === 0 && result.found > 0) {
      filesWithOnlyDuplicates.push({ file: fileName, count: result.found });
    }
  } catch (error) {
    filesWithErrors.push({ file: fileName, error: errorMsg });
  }
}

// Display report at end if any problematic files found
if (filesWithErrors.length > 0 || filesWithZeroEntries.length > 0 || filesWithOnlyDuplicates.length > 0) {
  console.log('\n=== File Status Report ===');
  // ... display each category
}
```

## Version History Location

See CHANGELOG.md for detailed version history. Current version defined in package.json.
