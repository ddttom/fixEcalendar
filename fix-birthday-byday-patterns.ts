#!/usr/bin/env ts-node
/**
 * Fix Corrupted Birthday BYDAY Patterns
 *
 * Issue: Some birthday entries use BYDAY (1st Thursday) instead of BYMONTHDAY (specific date)
 * This script converts BYDAY patterns to BYMONTHDAY for birthday/anniversary entries
 *
 * Example:
 *   Before: Birthday - Ian Asher (Jan 5, 2006) → FREQ=YEARLY;BYMONTH=1;BYDAY=1TH
 *   After:  Birthday - Ian Asher (Jan 5, 2006) → FREQ=YEARLY;BYMONTH=1;BYMONTHDAY=5
 *
 * NOTE: This creates entries marked as "Birthday (UNCERTAIN DATE)" to preserve data
 *       while flagging them for manual review.
 */

import Database from 'better-sqlite3';
import { join } from 'path';

interface BirthdayEntry {
  id: number;
  subject: string;
  start_time: number;
  recurrence_pattern: string;
}

async function fixBirthdayBydayPatterns() {
  const dbPath = join(process.cwd(), '.fixecalendar.db');
  console.log(`\nOpening database: ${dbPath}\n`);

  const db = new Database(dbPath);

  try {
    // Find birthday/anniversary entries with BYDAY patterns
    const query = `
      SELECT id, subject, start_time, recurrence_pattern
      FROM calendar_entries
      WHERE (subject LIKE '%Birthday%' OR subject LIKE '%Anniversary%')
        AND recurrence_pattern LIKE '%BYDAY=%'
      ORDER BY subject, start_time
    `;

    const entries = db.prepare(query).all() as BirthdayEntry[];

    console.log(`Found ${entries.length} birthday/anniversary entries with BYDAY patterns\n`);

    if (entries.length === 0) {
      console.log('✓ No corrupted birthday patterns found!\n');
      return;
    }

    console.log('⚠️  WARNING: Birthday entries should use fixed dates (BYMONTHDAY), not floating weekdays (BYDAY)\n');
    console.log('These entries will be updated to:\n');
    console.log('  1. Convert BYDAY to BYMONTHDAY using the actual start date');
    console.log('  2. Add "(UNCERTAIN DATE)" flag to subject for manual review\n');
    console.log('================\n');

    let fixedCount = 0;

    for (const entry of entries) {
      const startDate = new Date(entry.start_time);
      const month = startDate.getMonth() + 1;
      const day = startDate.getDate();

      // Convert BYDAY pattern to BYMONTHDAY
      const fixedPattern = entry.recurrence_pattern
        .replace(/BYDAY=[^;]+;?/, `BYMONTHDAY=${day};`)
        .replace(/;;/g, ';')
        .replace(/;$/, '');

      // Add warning to subject if not already present
      const fixedSubject = entry.subject.includes('UNCERTAIN')
        ? entry.subject
        : `${entry.subject} (UNCERTAIN DATE)`;

      // Update the database
      const updateStmt = db.prepare(`
        UPDATE calendar_entries
        SET recurrence_pattern = ?,
            subject = ?
        WHERE id = ?
      `);

      updateStmt.run(fixedPattern, fixedSubject, entry.id);
      fixedCount++;

      console.log(`✓ Fixed: ${entry.subject}`);
      console.log(`  Date: ${startDate.toDateString()} → Day ${day} of month`);
      console.log(`  Before: ${entry.recurrence_pattern}`);
      console.log(`  After:  ${fixedPattern}`);
      console.log(`  New Subject: ${fixedSubject}`);
      console.log();
    }

    console.log('================');
    console.log(`✓ Successfully fixed ${fixedCount} entries\n`);
    console.log('⚠️  IMPORTANT: Review entries marked "(UNCERTAIN DATE)" and correct if needed:\n');
    console.log('   sqlite3 .fixecalendar.db "SELECT id, subject, datetime(start_time/1000, \'unixepoch\') FROM calendar_entries WHERE subject LIKE \'%(UNCERTAIN DATE)%\'"');
    console.log('\nNext steps:');
    console.log('1. Review the uncertain entries above');
    console.log('2. Run: npx ts-node export-to-csv.ts');
    console.log('3. Run: npx ts-node export-to-ical.ts\n');
  } catch (error) {
    console.error('Error during fix:', error);
    throw error;
  } finally {
    db.close();
  }
}

// Run the fix
fixBirthdayBydayPatterns().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
