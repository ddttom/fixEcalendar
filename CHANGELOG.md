# Changelog

All notable changes to fixECalendar will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[1.2.2]: https://github.com/ddttom/fixEcalendar/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/ddttom/fixEcalendar/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/ddttom/fixEcalendar/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/ddttom/fixEcalendar/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ddttom/fixEcalendar/releases/tag/v1.0.0
