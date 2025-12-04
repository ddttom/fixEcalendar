#!/usr/bin/env ts-node
/**
 * Database Cleanup Script: Fix Missing BYMONTH in Yearly BYDAY Recurrence Patterns
 *
 * Issue: Events with FREQ=YEARLY;BYDAY=-1SA (last Saturday) fail to import into Google Calendar
 * Root Cause: Missing BYMONTH parameter makes RFC 5545 interpret it as "last Saturday of the year"
 *              instead of "last Saturday of the specific month"
 *
 * This script:
 * 1. Finds all calendar entries with FREQ=YEARLY + BYDAY but missing BYMONTH
 * 2. Extracts the month from the event's start_time
 * 3. Inserts BYMONTH parameter into the recurrence pattern
 * 4. Updates the database
 *
 * Usage: npx ts-node cleanup-yearly-byday.ts
 */

import Database from 'better-sqlite3';
import { join } from 'path';

interface AffectedEntry {
  id: number;
  subject: string;
  start_time: number;
  recurrence_pattern: string;
}

async function cleanupYearlyByday() {
  const dbPath = join(process.cwd(), '.fixecalendar.db');
  console.log(`\nOpening database: ${dbPath}\n`);

  const db = new Database(dbPath);

  try {
    // Find all entries with YEARLY + BYDAY but no BYMONTH
    const query = `
      SELECT id, subject, start_time, recurrence_pattern
      FROM calendar_entries
      WHERE recurrence_pattern LIKE '%FREQ=YEARLY%'
        AND recurrence_pattern LIKE '%BYDAY=%'
        AND recurrence_pattern NOT LIKE '%BYMONTH=%'
      ORDER BY subject, start_time
    `;

    const entries = db.prepare(query).all() as AffectedEntry[];

    console.log(`Found ${entries.length} entries with YEARLY BYDAY missing BYMONTH\n`);

    if (entries.length === 0) {
      console.log('✓ No entries need fixing. Database is clean!\n');
      return;
    }

    console.log('Affected entries:');
    console.log('================\n');

    let fixedCount = 0;

    for (const entry of entries) {
      // Extract month from start_time (milliseconds since epoch)
      const startDate = new Date(entry.start_time);
      const month = startDate.getMonth() + 1; // getMonth() returns 0-11, we need 1-12

      // Insert BYMONTH after FREQ=YEARLY
      const fixedPattern = entry.recurrence_pattern.replace(
        /FREQ=YEARLY;?/,
        `FREQ=YEARLY;BYMONTH=${month};`
      );

      // Update the database
      const updateStmt = db.prepare(`
        UPDATE calendar_entries
        SET recurrence_pattern = ?
        WHERE id = ?
      `);

      updateStmt.run(fixedPattern, entry.id);
      fixedCount++;

      // Display the fix
      console.log(`✓ Fixed: ${entry.subject}`);
      console.log(`  Start Date: ${startDate.toDateString()} (Month ${month})`);
      console.log(`  Before: ${entry.recurrence_pattern}`);
      console.log(`  After:  ${fixedPattern}`);
      console.log();
    }

    console.log('================');
    console.log(`✓ Successfully fixed ${fixedCount} entries\n`);
    console.log('Next steps:');
    console.log('1. Run: npx ts-node export-to-csv.ts');
    console.log('2. Run: npx ts-node export-to-ical.ts');
    console.log('3. Verify: grep -A 5 "York Residents First Weekend" calendar-part-*.ics');
    console.log('4. Import the fixed ICS files into Google Calendar\n');
  } catch (error) {
    console.error('Error during cleanup:', error);
    throw error;
  } finally {
    db.close();
  }
}

// Run the cleanup
cleanupYearlyByday().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
