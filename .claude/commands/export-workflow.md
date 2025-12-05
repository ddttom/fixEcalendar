---
description: Run complete export workflow (Database → CSV → ICS)
---

Run the complete export workflow to generate both CSV and ICS files from the database:

1. Export database to CSV format (output/calendar-export.csv)
2. Convert CSV to ICS format (output/calendar-export.ics)

This workflow creates two output files in the `output/` folder:
- **output/calendar-export.csv**: For review/editing in Excel/Google Sheets
- **output/calendar-export.ics**: For importing into calendar applications

Execute the following commands:

```bash
# Step 1: Export database to CSV
npx ts-node src/scripts/export-to-csv.ts

# Step 2: Convert CSV to ICS
npx ts-node src/scripts/export-to-ical.ts
```

Both commands run sequentially - the ICS export depends on the CSV file being created first.
