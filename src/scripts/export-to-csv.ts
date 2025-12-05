#!/usr/bin/env ts-node

/**
 * Export calendar entries from database to CSV
 */

import * as fs from 'fs';
import * as path from 'path';
import Database from 'better-sqlite3';
import { formatDescription } from '../utils/text-formatter';

function escapeCSV(value: any): string {
  if (value === null || value === undefined) {
    return '""';
  }

  const str = String(value);

  // Always wrap in quotes for Excel compatibility
  // Escape quotes by doubling them and replace newlines with literal \n
  return `"${str.replace(/"/g, '""').replace(/\r?\n/g, '\\n')}"`;
}

function standardizeBirthdaySubject(subject: string, startDate: Date, recurrencePattern: string | null): string {
  // Check if this is a birthday/anniversary
  const subjectLower = subject.toLowerCase();
  if (!subjectLower.includes('birthday') && !subjectLower.includes('anniversary')) {
    return subject;
  }

  // Extract day and month from recurrence pattern BYMONTHDAY
  if (!recurrencePattern) {
    return subject;
  }

  const byMonthDayMatch = recurrencePattern.match(/BYMONTHDAY=(\d+)/);
  if (!byMonthDayMatch) {
    return subject;
  }

  const day = parseInt(byMonthDayMatch[1]);
  const month = startDate.getMonth() + 1; // 0-indexed, so add 1
  const year = startDate.getFullYear();

  // Format as dd/mm/yyyy
  const dateStr = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;

  // Remove any existing date pattern in parentheses (including text dates like "5th May 1979")
  const cleanSubject = subject
    .replace(/\s*\([^)]*\d{4}[^)]*\)/g, '') // Remove any parentheses containing a 4-digit year
    .replace(/\s*\(\s*\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]?\d{2,4}\s*\)/g, '') // Remove (dd/mm/yyyy) or similar
    .replace(/\s*\(\s*\d{4}\s*\)/g, '') // Remove (yyyy)
    .replace(/\s*\(\s*\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2}\s*\)/g, '') // Remove (dd/mm/yy)
    .trim();

  return `${cleanSubject} (${dateStr})`;
}

function formatDate(date: Date): string {
  return date.toISOString();
}

async function exportToCSV() {
  const dbPath = '.fixecalendar.db';

  if (!fs.existsSync(dbPath)) {
    console.error('Database not found. Please process PST files first.');
    process.exit(1);
  }

  console.log('Opening database...');
  const db = new Database(dbPath, { readonly: true });

  console.log('Retrieving entries...');
  const entries = db.prepare('SELECT * FROM calendar_entries ORDER BY start_time ASC').all() as any[];

  console.log(`Found ${entries.length} entries`);

  // CSV headers
  const headers = [
    'Subject',
    'Start Date',
    'Start Time',
    'End Date',
    'End Time',
    'Location',
    'Description',
    'Organizer',
    'All Day',
    'Importance',
    'Busy Status',
    'Sensitivity',
    'Is Recurring',
    'Recurrence Pattern',
    'Reminder (minutes)'
  ];

  // Build CSV content
  const rows: string[] = [];
  rows.push(headers.map(escapeCSV).join(','));

  // Track unique entries to avoid duplicates after date normalization
  const seen = new Set<string>();

  for (const entry of entries) {
    let startDate = new Date(entry.start_time);
    let endDate = new Date(entry.end_time);
    let isAllDayEvent = entry.is_all_day;

    // Special handling for birthdays/anniversaries - force them to be all-day events
    const subject = (entry.subject || '').toLowerCase();
    const isBirthdayOrAnniversary = subject.includes('birthday') || subject.includes('anniversary');

    if (isBirthdayOrAnniversary && entry.recurrence_pattern) {
      // Extract day from BYMONTHDAY in recurrence pattern
      const byMonthDayMatch = entry.recurrence_pattern.match(/BYMONTHDAY=(\d+)/);
      if (byMonthDayMatch) {
        const correctDay = parseInt(byMonthDayMatch[1]);
        // Create date with the correct day from recurrence pattern
        // Use the start date's year and month to determine the correct date
        const baseDate = new Date(entry.start_time);
        const year = baseDate.getFullYear();
        const month = baseDate.getMonth();

        // Adjust month if the stored date is before the correct day (due to UTC offset)
        const adjustedMonth = baseDate.getDate() < correctDay ? month : month;

        startDate = new Date(year, adjustedMonth, correctDay);
        endDate = new Date(year, adjustedMonth, correctDay); // Same day for all-day events
        isAllDayEvent = true; // Force as all-day event
      }
    }

    // Handle all-day events - normalize to local date midnight
    if (isAllDayEvent) {
      // Create new dates at local midnight (remove time component)
      startDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      endDate = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
    }

    // Format dates properly - use local date components for all-day events, ISO for timed events
    const startDateStr = isAllDayEvent
      ? `${startDate.getFullYear()}-${String(startDate.getMonth() + 1).padStart(2, '0')}-${String(startDate.getDate()).padStart(2, '0')}`
      : startDate.toISOString().split('T')[0];
    const startTimeStr = isAllDayEvent
      ? 'All Day'
      : startDate.toISOString().split('T')[1].split('.')[0];
    const endDateStr = isAllDayEvent
      ? `${endDate.getFullYear()}-${String(endDate.getMonth() + 1).padStart(2, '0')}-${String(endDate.getDate()).padStart(2, '0')}`
      : endDate.toISOString().split('T')[0];
    const endTimeStr = isAllDayEvent
      ? 'All Day'
      : endDate.toISOString().split('T')[1].split('.')[0];

    // Standardize birthday/anniversary subject format
    const standardizedSubject = standardizeBirthdaySubject(
      entry.subject || '',
      new Date(entry.start_time),
      entry.recurrence_pattern
    );

    // Create unique key for deduplication (subject + start date + start time)
    const uniqueKey = `${standardizedSubject}|${startDateStr}|${startTimeStr}`;

    // Skip if we've already seen this exact entry
    if (seen.has(uniqueKey)) {
      continue;
    }
    seen.add(uniqueKey);

    const row = [
      standardizedSubject,
      startDateStr,
      startTimeStr,
      endDateStr,
      endTimeStr,
      entry.location || '',
      formatDescription(entry.description),
      entry.organizer || '',
      isAllDayEvent ? 'Yes' : 'No',
      entry.importance || 'normal',
      entry.busy_status || 'busy',
      entry.sensitivity || 'normal',
      entry.is_recurring ? 'Yes' : 'No',
      entry.recurrence_pattern || '',
      entry.reminder || ''
    ];

    rows.push(row.map(escapeCSV).join(','));
  }

  db.close();

  const csvContent = rows.join('\n');

  const outputPath = path.join('output', 'calendar-export.csv');

  // Ensure output directory exists
  const outputDir = path.dirname(outputPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  fs.writeFileSync(outputPath, csvContent, 'utf-8');

  const uniqueEntries = rows.length - 1; // Subtract header row
  const duplicatesRemoved = entries.length - uniqueEntries;

  console.log(`\n✓ Successfully exported ${uniqueEntries} unique entries to ${outputPath}`);
  if (duplicatesRemoved > 0) {
    console.log(`  (${duplicatesRemoved} duplicates removed after date normalization)`);
  }
  console.log(`\nYou can open this file in Excel, Google Sheets, or any spreadsheet application.`);
}

exportToCSV().catch(error => {
  console.error('Error:', error.message);
  process.exit(1);
});
