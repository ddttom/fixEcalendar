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
- Recurrence pattern parsing from binary RecurrencePattern class

### Text Formatting
Located in `src/utils/text-formatter.ts`:
- `formatDescription()`: Trims whitespace, normalizes line breaks, truncates to 79 chars
- Applied consistently across all export formats
- Prevents unwieldy text in exports

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

### ICS Export from CSV (`export-to-ical.ts`)
Standalone script that converts CSV export to iCalendar (ICS) format.
- Reads from `calendar-export.csv`
- Parses CSV with proper quoted field handling
- Converts to CalendarEntry objects
- Uses ICalConverter to generate RFC 5545 compliant `calendar-export.ics`
- Preserves all properties including attendees, recurrence, reminders
- Reuses existing PropertyMapper for consistent formatting
- **Includes automatic cleanup**: Removes invalid `INTERVAL=12` from yearly recurrence rules for Google Calendar compatibility

**Workflow:**
```
Database → export-to-csv.ts → calendar-export.csv → export-to-ical.ts → calendar-export.ics
```

This two-step process allows for CSV review/editing before ICS generation.

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

## Version History Location

See CHANGELOG.md for detailed version history. Current version defined in package.json.
