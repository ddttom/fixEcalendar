#!/usr/bin/env ts-node
/**
 * Database cleanup script for suspicious recurrence patterns
 *
 * This script finds all calendar entries with UNTIL=2100 recurrence patterns
 * and applies validation logic to cap them at reasonable durations based on
 * frequency (daily: 5 years, weekly: 10 years, monthly: 20 years, yearly: 100 years).
 *
 * Usage: npx ts-node cleanup-suspicious-recurrence.ts
 */

import Database from 'better-sqlite3';
import { RecurrenceValidator } from './src/utils/recurrence-validator';
import * as fs from 'fs';

interface CleanupResult {
  id: number;
  subject: string;
  startDate: string;
  yearsSpan: number;
  oldPattern: string;
  newPattern: string | null;
  action: 'STRIPPED' | 'CAPPED';
  reason: string;
}

function escapeCSV(value: any): string {
  const str = String(value || '');
  // Escape quotes by doubling them and replace newlines with literal \n
  return `"${str.replace(/"/g, '""').replace(/\r?\n/g, '\\n')}"`;
}

function generateCSVReport(results: CleanupResult[]): string {
  const headers = [
    'ID',
    'Subject',
    'Start Date',
    'Years Span',
    'Action',
    'Old Pattern',
    'New Pattern',
    'Reason',
  ];

  const rows = results.map((r) => [
    r.id,
    r.subject,
    r.startDate,
    r.yearsSpan,
    r.action,
    r.oldPattern,
    r.newPattern || 'STRIPPED',
    r.reason,
  ]);

  const csvLines = [headers.map(escapeCSV).join(',')];
  for (const row of rows) {
    csvLines.push(row.map(escapeCSV).join(','));
  }

  return csvLines.join('\n');
}

function cleanupSuspiciousRecurrence() {
  const dbPath = '.fixecalendar.db';

  if (!fs.existsSync(dbPath)) {
    console.error(`\n❌ Database not found: ${dbPath}`);
    console.error('Please run the PST import first to create the database.\n');
    process.exit(1);
  }

  const db = new Database(dbPath);

  console.log('\n=== Recurrence Pattern Cleanup Script ===\n');
  console.log('Scanning database for suspicious recurrence patterns...\n');

  // Find all entries with UNTIL=2100
  const entries = db
    .prepare(
      `
    SELECT id, subject, start_time, recurrence_pattern
    FROM calendar_entries
    WHERE recurrence_pattern LIKE '%UNTIL=2100%'
  `
    )
    .all() as Array<{
    id: number;
    subject: string;
    start_time: number;
    recurrence_pattern: string;
  }>;

  if (entries.length === 0) {
    console.log('✓ No suspicious recurrence patterns found.');
    console.log('  All entries have reasonable recurrence durations.\n');
    db.close();
    return;
  }

  console.log(`Found ${entries.length} entries with UNTIL=2100\n`);

  const results: CleanupResult[] = [];
  const updateStmt = db.prepare('UPDATE calendar_entries SET recurrence_pattern = ? WHERE id = ?');

  let cappedCount = 0;
  let strippedCount = 0;

  for (const entry of entries) {
    const startDate = new Date(entry.start_time);
    const yearsSpan = 2100 - startDate.getFullYear();

    // Apply validation
    const validated = RecurrenceValidator.validateAndCap(
      entry.recurrence_pattern,
      startDate,
      entry.subject
    );

    if (validated.modified) {
      // Update database
      updateStmt.run(validated.newPattern || null, entry.id);

      results.push({
        id: entry.id,
        subject: entry.subject,
        startDate: startDate.toISOString().split('T')[0],
        yearsSpan,
        oldPattern: entry.recurrence_pattern,
        newPattern: validated.newPattern,
        action: validated.newPattern ? 'CAPPED' : 'STRIPPED',
        reason: validated.reason,
      });

      if (validated.newPattern) {
        cappedCount++;
        console.log(
          `✓ CAPPED: ${entry.subject} (${yearsSpan} years → ${validated.newYearsSpan} years)`
        );
      } else {
        strippedCount++;
        console.log(`✓ STRIPPED: ${entry.subject} (${yearsSpan} years - unreasonable)`);
      }
    }
  }

  db.close();

  // Generate CSV report
  const reportPath = 'recurrence-cleanup-report.csv';
  const csv = generateCSVReport(results);
  fs.writeFileSync(reportPath, csv);

  // Summary
  console.log(`\n=== Cleanup Summary ===\n`);
  console.log(`Total entries scanned: ${entries.length}`);
  console.log(`Patterns modified: ${results.length}`);
  console.log(`  - Capped: ${cappedCount}`);
  console.log(`  - Stripped: ${strippedCount}`);
  console.log(`\n✓ Report saved to: ${reportPath}`);

  // Recommendations
  console.log(`\n=== Next Steps ===\n`);
  console.log(`1. Review the cleanup report:`);
  console.log(`   cat ${reportPath}\n`);
  console.log(`2. Re-export to CSV with fixed patterns:`);
  console.log(`   npx ts-node export-to-csv.ts\n`);
  console.log(`3. Generate ICS files:`);
  console.log(`   npx ts-node export-to-ical.ts\n`);
  console.log(`4. Import the fixed ICS files to your calendar application\n`);

  // Breakdown by frequency
  const byFrequency: { [key: string]: number } = {};
  for (const result of results) {
    const freqMatch = result.oldPattern.match(/FREQ=(\w+)/);
    const freq = freqMatch ? freqMatch[1] : 'UNKNOWN';
    byFrequency[freq] = (byFrequency[freq] || 0) + 1;
  }

  if (Object.keys(byFrequency).length > 0) {
    console.log(`=== Breakdown by Frequency ===\n`);
    for (const [freq, count] of Object.entries(byFrequency)) {
      console.log(`  ${freq}: ${count} entries`);
    }
    console.log('');
  }
}

// Run the cleanup
try {
  cleanupSuspiciousRecurrence();
} catch (error) {
  console.error('\n❌ Error during cleanup:', (error as Error).message);
  console.error((error as Error).stack);
  process.exit(1);
}
