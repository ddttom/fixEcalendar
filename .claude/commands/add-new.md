Add a new PST file to the database with automatic deduplication.

This command processes a PST file and adds all calendar entries to the `.fixecalendar.db` database. The database automatically prevents duplicates using SHA256 hash deduplication.

## Usage

Ask the user for the PST file path, then run:

```bash
node dist/index-with-db.js "<pst-file-path>" --verbose
```

Replace `<pst-file-path>` with the actual path to the PST file.

## What This Does

1. **Opens PST file** - Validates and opens the Microsoft Outlook PST file
2. **Extracts calendar entries** - Finds the calendar folder and extracts all appointments
3. **Deduplicates** - Automatically skips entries that already exist in the database (based on hash of uid, subject, start/end time, location, organizer)
4. **Stores in database** - Adds new entries to `.fixecalendar.db`
5. **Shows summary** - Displays how many entries were found, added, and skipped

## Expected Output

The command will display:
- Number of calendar entries found
- Number of new entries added to database
- Number of duplicates skipped
- Processing summary with database statistics

## Notes

- **Automatic deduplication**: The database uses SHA256 hashing to detect duplicates, so running this command multiple times with the same PST file will not create duplicates
- **Verbose mode**: The `--verbose` flag provides detailed logging of the extraction process
- **Database location**: By default uses `.fixecalendar.db` in the current directory
- **No export**: This command only adds to the database. To export to ICS format, use a separate export command

## Error Handling

If the command fails, check:
- PST file path is correct and file exists
- PST file is a valid Microsoft Outlook PST file
- PST file contains a calendar folder
- You have read permissions for the PST file
- The `dist/` directory exists (run `npm run build` if needed)
