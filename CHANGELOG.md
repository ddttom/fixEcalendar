# Changelog

All notable changes to fixECalendar will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- File Status Report: Shows summary of problematic PST files at end of processing
  - Files with errors (unreadable, corrupted, or no calendar folders)
  - Files with zero calendar entries
  - Files where all entries were duplicates
- **Overlapping Event Merge**: Automatic merging of overlapping events with same description
  - New script `merge-overlapping-events.ts` for manual/automatic merging
  - Subject normalization removes time prefixes (e.g., "09:30", "11:05")
  - Merges events with earliest start and latest end time
  - Selects best description from multiple entries
  - Dry-run support for safe preview
- **Automatic Merge Integration**: Runs automatically after PST import
  - `--skip-merge` CLI option to disable automatic merging
  - Non-fatal failure handling with manual fallback
- **Birthday Pattern Fix Script**: `fix-birthday-byday-patterns.ts`
  - Converts corrupted BYDAY patterns to BYMONTHDAY for birthdays/anniversaries
  - Adds "(UNCERTAIN DATE)" flag for manual review
  - Preserves all data per data preservation policy
- **Yearly BYDAY Cleanup Script**: `cleanup-yearly-byday.ts`
  - One-time database repair for missing BYMONTH parameters
  - Fixes RFC 5545 compliance for yearly recurring events
- **Data Preservation Policy**: Added to CLAUDE.md
  - Never delete entries without explicit user request
  - Convert, mark, and preserve suspicious data instead

### Changed
- **Breaking Change**: Now processes ALL calendar folders found in a PST file, not just the largest one
- Multiple calendar folders in a single PST are now all processed automatically
- Clear folder-by-folder progress indication when processing multiple folders
- Total summary shown when multiple folders are processed
- **RFC 5545 Enhancement**: Yearly BYDAY patterns now include BYMONTH parameter
  - PST import adds BYMONTH during extraction
  - CSV import includes cleanup for legacy data

### Fixed
- Calendar folder detection now uses Microsoft's PR_CONTAINER_CLASS property (containerClass)
- Folders like "Calendar (This computer only)" are now correctly recognized as calendar folders
- Parser now finds and processes ALL calendar folders in PST file, including nested folders
- Enhanced logging shows displayName, containerClass, and entry count for troubleshooting
- Fallback to name-based detection ensures backward compatibility with existing PST files
- **Critical**: Missing BYMONTH in yearly BYDAY patterns causing Google Calendar import failures
  - "York Residents First Weekend" and similar events now import correctly
  - Without BYMONTH: `FREQ=YEARLY;BYDAY=-1SA` means last Saturday of YEAR (December)
  - With BYMONTH: `FREQ=YEARLY;BYMONTH=1;BYDAY=-1SA` means last Saturday of JANUARY
  - Fixed 7 entries with yearly BYDAY patterns
- Corrupted birthday patterns using BYDAY (floating weekday) instead of BYMONTHDAY (specific date)
  - Converted to BYMONTHDAY using actual start date
  - Marked with "(UNCERTAIN DATE)" for manual verification
- Overlapping events creating calendar clutter
  - Events like "09:30 Aquafit", "11:05 Aquafit" now merged into single "Aquafit" event
  - Merged 7 groups of overlapping events in typical calendar

## [1.2.5] - 2025-12-04

### Added
- **Recurrence Pattern Validation System**: Comprehensive validation to prevent absurd recurring schedules
  - New module `src/utils/recurrence-validator.ts` with configurable caps per frequency
  - DAILY: Maximum 5 years (unless infinite/COUNT-based)
  - WEEKLY: Maximum 10 years
  - MONTHLY: Maximum 20 years
  - YEARLY: Maximum 100 years (birthdays/anniversaries)
  - Single occurrence: Strips recurrence if `occurrenceCount === 1`
- **Prefer COUNT over corrupted UNTIL**: When UNTIL date is corrupted (year < 1900) AND occurrenceCount exists (1-50), uses COUNT parameter
  - Provides exact occurrence counts for finite recurring events (e.g., 4-day course = `COUNT=4`)
  - More accurate than capping corrupted UNTIL dates to arbitrary future years
  - 202 entries now use COUNT-based patterns instead of capped UNTIL dates
- Database cleanup script `cleanup-suspicious-recurrence.ts` for existing entries
- Configuration constants in `src/config/constants.ts` for customizable thresholds

### Changed
- PST import now validates all recurrence patterns automatically
- CSV import (`export-to-ical.ts`) includes defense-in-depth cleanup for UNTIL=2100 patterns

### Fixed
- **Critical**: Events with corrupted recurrence patterns creating 70-85 year schedules (e.g., "Toms Online Photo Course" repeating daily for 76 years = 27,740 occurrences)
- Fixed 72 entries with UNTIL=2100 spanning 70-85 years:
  - 11 DAILY entries: Capped to 5 years or converted to COUNT
  - 45 WEEKLY entries: Capped to 10 years
  - 13 MONTHLY entries: Capped to 20 years
  - 3 YEARLY entries: Kept as-is (legitimate birthdays)
- Example: "Toms Online Photo Course" changed from 76 years (27,740 occurrences) to `COUNT=4` (4 occurrences)

## [1.2.4] - 2025-12-04

### Added
- Script `sanitize-recurrence-dates.ts` to fix corrupted UNTIL dates in existing database entries
- Automatic file splitting for Google Calendar compatibility (499-event chunks)
- Historical date handling for recurring birthdays/anniversaries before 1970

### Changed
- Birthday/anniversary events before 1970 now use modern anchor date (2020) for Google Calendar compatibility
- Original birth year preserved in subject line, actual event dates unchanged

### Fixed
- **Critical**: CSV export now properly escapes newlines in descriptions
  - Multi-line descriptions converted to literal `\n` in CSV to prevent row splitting
  - CSV import correctly unescapes `\n` back to actual newlines for ICS generation
  - Result: 100% data integrity in CSV → ICS conversion (previously lost ~50% of entries)
  - Impact: Before fix: 2,470/4,886 entries converted. After fix: 4,886/4,886 entries converted
- **Critical**: Corrupted recurrence UNTIL dates now sanitized (Microsoft Outlook bug)
  - Recurrence patterns with UNTIL dates before year 1900 projected to year 2100
  - Fixed 71 corrupted entries with `UNTIL=16001231` → `UNTIL=21001231`
- **Critical**: All-day events now RFC 5545 compliant for Google Calendar import
  - Birthday/anniversary events now have proper exclusive DTEND dates
  - DTEND for all-day events is now the day after the event (exclusive end date per RFC 5545)
- **Critical**: Recurring events now RFC 5545 compliant (UNTIL time component and daily intervals)
  - UNTIL values now include time component (e.g., `UNTIL=21001231T235959Z`)
  - Daily recurrence intervals converted from minutes to days (MS-OXOCAL spec compliance)
  - Fixed 72 recurring events with date-only UNTIL, converted `INTERVAL=1440` (minutes) to `INTERVAL=1` (day)
- **Critical**: Yearly recurring events now RFC 5545 compliant with BYMONTH parameter
  - Added BYMONTH parameter to yearly events with BYMONTHDAY
  - Fixed 465 yearly recurring events that Google Calendar was rejecting
- Description field now filters out HTML/CSS junk from malformed Outlook extractions
  - Removes patterns like "false\nfalse\nfalse", "EN-GB", "X-NONE", "/* Style Definitions */", "table.MsoNormal"
  - Cleaned 1,832 corrupted descriptions
- Export-to-ical.ts now automatically splits large calendars into 499-event chunks
  - Files with 499 events import successfully into Google Calendar (tested and confirmed)
- All recurring birthdays now import successfully while preserving historical event dates (2000-2004)
- ICS files now import cleanly into Google Calendar, Apple Calendar, and all RFC 5545 compliant applications

## [1.2.3] - 2025-12-04

### Added
- **Intelligent Date Recovery & Data Quality**: Automatically recovers valid appointments and filters corrupted data
  - Uses duration field to calculate missing start/end times
  - Falls back to alternative date fields (recurrenceBase, creationTime, etc.)
  - Fixes reversed dates (endTime before startTime)
  - Handles zero-duration appointments intelligently (1 hour for regular, 24 hours for birthdays/anniversaries)
  - Automatically converts birthdays/anniversaries/holidays to proper all-day events
  - Discards corrupted entries with no subject to prevent database bloat
  - Result: 60% reduction in database bloat while recovering more valid entries

### Fixed
- **Critical**: Calendar folder detection now handles non-standard folder names and nested structures
- Uses Microsoft's PR_CONTAINER_CLASS property (`IPF.Appointment`) for reliable folder detection
- Folders like "Calendar (This computer only)" are now correctly recognized
- Parser now finds ALL calendar folders, including nested folders
- Automatically selects folder with most entries when multiple folders exist
- Enhanced logging shows displayName, containerClass, and entry count
- Successfully processes PST files with nested calendar folder structures
- Backward compatible with existing name-based detection

## [1.2.2] - 2025-12-04

### Added
- CSV to ICS conversion script (`export-to-ical.ts`)
- Ability to convert `calendar-export.csv` to `calendar-export.ics`
- Full RFC 5545 compliant iCalendar generation from CSV data
- Proper CSV parsing with quoted field handling
- Date/time conversion from CSV format to CalendarEntry objects
- Deduplication in CSV export to prevent duplicates after date normalization
- Two-step export workflow documentation (Database → CSV → ICS)
- Visual workflow diagrams in README showing complete export pipeline
- Automatic deduplication reporting with statistics
- Basic test suite to satisfy Jest test runner

### Changed
- `export-to-ical.ts` now reads from CSV file instead of database
- Simplified workflow: export to CSV for analysis, then convert to ICS for import
- README restructured with prominent workflow sections at top
- Export workflows now clearly documented in Quick Start section
- Updated `.gitignore` to allow `package-lock.json` for CI/CD
- CI/CD now tests only Node.js 18.x and 20.x (removed 16.x due to dependency requirements)
- Updated `package.json` engines field to require Node.js 18.18.0 or higher
- Updated README prerequisites to specify Node.js 18.18.0 or higher

### Fixed
- Duplicate entries in CSV export caused by birthday/anniversary date normalization
- CSV export now tracks unique entries by subject + date + time
- Reports number of duplicates removed during export (e.g., "49 duplicates removed")
- GitHub Actions CI/CD failures due to missing `package-lock.json`
- Dependency caching now working in all CI workflows
- Node 16.x compatibility issues with modern ESLint and better-sqlite3
- Missing test files causing Jest to fail in CI
- Unused `fs` imports in index files causing ESLint errors
- **Critical: Invalid RRULE format causing silent Google Calendar import failures**
  - Yearly recurrence now correctly uses `FREQ=YEARLY` without `INTERVAL=12`
  - Outlook stores yearly events as 12-month intervals, now properly converted to iCalendar standard
  - Added automatic cleanup in `export-to-ical.ts` for existing CSV data with invalid intervals
  - ICS files now import successfully into Google Calendar and other RFC 5545 compliant applications

## [1.2.1] - 2025-12-04

### Added
- Created centralized text formatting utility (`src/utils/text-formatter.ts`)
- `formatDescription()` function for consistent description formatting across exports

### Changed
- Description fields now automatically cleaned and formatted in all exports
- Whitespace trimmed from both ends of descriptions
- Multiple consecutive line breaks (3 or more) reduced to exactly 2 for readability
- Description fields truncated to 79 characters maximum

### Fixed
- Excessive whitespace in description fields across CSV and ICS exports
- Unwieldy description text that made exports difficult to read

## [1.2.0] - 2025-12-03

### Added
- CSV export functionality via `export-to-csv.ts` script
- Comprehensive RRULE recurrence pattern support using pst-extractor's RecurrencePattern class
- Support for DAILY, WEEKLY, MONTHLY, YEARLY frequencies with ordinals, intervals, and counts
- Birthday/anniversary subject standardization to `(dd/mm/yyyy)` format

### Changed
- All-day events now use proper `VALUE=DATE` format in iCal exports
- CSV export properly normalizes all-day event dates to remove UTC timezone offsets

### Fixed
- All-day event timezone handling - birthdays/anniversaries now display correctly as single-day events (no more 2-day spans)
- CSV export now properly quotes all fields for correct Excel import
- Leading zeros preserved in CSV exports
- Dates not auto-formatted incorrectly in Excel
- Multi-line descriptions and special characters handled correctly
- Commas within field data don't break CSV structure

## [1.1.0] - 2025-12-03

### Added
- SQLite database for intermediate storage of calendar entries
- Automatic deduplication of calendar entries based on UID, subject, times, location, and organizer
- `--database` option to specify custom database location
- `--clear-db` option to clear database before processing
- `--db-stats` option to view database statistics
- `--no-dedupe` option to disable deduplication (not recommended)
- Processing statistics and logging for large file handling

### Changed
- Optimized for very large PST files (6GB+)
- Improved memory efficiency - uses ~500MB-1GB instead of loading entire file
- Processing speed increased to ~1,000-2,000 entries/second

### Fixed
- Memory issues when processing large PST files
- Duplicate entries when re-processing same files
- Inability to incrementally add PST files over time

## [1.0.0] - 2025-12-03

### Added
- Initial release
- PST to iCal (.ics) conversion
- Batch processing of multiple PST files
- Glob pattern support for file discovery
- Merge mode to combine multiple calendars
- Full property mapping (attendees, locations, descriptions, reminders, etc.)
- Support for all-day events
- Timezone preservation
- Command-line interface with comprehensive options
- TypeScript implementation with type safety
- Support for both ANSI and Unicode PST formats
- `--include-private` flag to include private/confidential appointments
- `--no-recurring` flag to exclude recurring appointments
- Date range filtering with `--start-date` and `--end-date`
- Verbose logging with `--verbose` flag
- Custom calendar naming with `--name` option
- Timezone setting with `--timezone` option

### Dependencies
- pst-extractor - Pure JavaScript PST file parser
- ical-generator - iCalendar file generation
- commander - Command-line interface framework
- fast-glob - Fast file pattern matching
- better-sqlite3 - Fast, synchronous SQLite3 database

[1.2.5]: https://github.com/ddttom/fixEcalendar/compare/v1.2.4...v1.2.5
[1.2.4]: https://github.com/ddttom/fixEcalendar/compare/v1.2.3...v1.2.4
[1.2.3]: https://github.com/ddttom/fixEcalendar/compare/v1.2.2...v1.2.3
[1.2.2]: https://github.com/ddttom/fixEcalendar/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/ddttom/fixEcalendar/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/ddttom/fixEcalendar/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/ddttom/fixEcalendar/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ddttom/fixEcalendar/releases/tag/v1.0.0
