#!/usr/bin/env ts-node

/**
 * Sanitize corrupted recurrence UNTIL dates in the database
 *
 * This script fixes a known Microsoft Outlook bug where recurrence patterns
 * sometimes have UNTIL dates set to year 1600 (UNTIL=16001231).
 *
 * The script projects these corrupted dates to year 2100 for better
 * calendar application compatibility.
 */

import Database from 'better-sqlite3';

function sanitizeRecurrenceDates() {
  const dbPath = '.fixecalendar.db';

  console.log('Opening database...');
  const db = new Database(dbPath);

  // Find all entries with corrupted UNTIL dates (year < 1900)
  console.log('Searching for entries with corrupted recurrence dates...');

  const corrupted = db.prepare(`
    SELECT id, subject, recurrence_pattern
    FROM calendar_entries
    WHERE recurrence_pattern LIKE '%UNTIL=%'
  `).all() as Array<{ id: number; subject: string; recurrence_pattern: string }>;

  console.log(`Found ${corrupted.length} entries with UNTIL clauses`);

  let fixedCount = 0;
  const updateStmt = db.prepare('UPDATE calendar_entries SET recurrence_pattern = ? WHERE id = ?');

  for (const entry of corrupted) {
    // Extract the UNTIL date
    const untilMatch = entry.recurrence_pattern.match(/UNTIL=(\d{8})/);
    if (!untilMatch) continue;

    const untilDate = untilMatch[1];
    const year = parseInt(untilDate.substring(0, 4));

    // Check if year is before 1900 (corrupted)
    if (year < 1900) {
      // Replace the year with 2100, keep month and day
      const monthDay = untilDate.substring(4); // Get MMDD
      const newUntilDate = `2100${monthDay}`;
      const newPattern = entry.recurrence_pattern.replace(/UNTIL=\d{8}/, `UNTIL=${newUntilDate}`);

      updateStmt.run(newPattern, entry.id);
      fixedCount++;
      console.log(`Fixed: "${entry.subject}" - UNTIL=${untilDate} → UNTIL=${newUntilDate}`);
    }
  }

  db.close();

  console.log(`\n✓ Fixed ${fixedCount} corrupted recurrence dates (projected to year 2100)`);
  console.log(`  ${corrupted.length - fixedCount} entries had valid dates (no changes needed)`);

  if (fixedCount > 0) {
    console.log('\nRun export scripts to update CSV/ICS files:');
    console.log('  npx ts-node export-to-csv.ts');
    console.log('  npx ts-node export-to-ical.ts');
  }
}

sanitizeRecurrenceDates();
