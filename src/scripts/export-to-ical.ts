#!/usr/bin/env ts-node

/**
 * Export calendar entries from CSV to iCalendar (ICS) format
 */

import * as fs from 'fs';
import { CalendarEntry } from '../parser/types';
import { ICalConverter } from '../converter/ical-converter';

/**
 * Parse CSV line handling quoted fields
 */
function parseCSVLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        // Escaped quote
        current += '"';
        i++; // Skip next quote
      } else {
        // Toggle quote state
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // Field separator
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }

  // Add last field
  result.push(current);

  return result;
}

/**
 * Convert CSV row to CalendarEntry object
 */
function rowToEntry(fields: string[]): CalendarEntry {
  // CSV fields: Subject, Start Date, Start Time, End Date, End Time, Location, Description,
  // Organizer, All Day, Importance, Busy Status, Sensitivity, Is Recurring, Recurrence Pattern, Reminder (minutes)

  const subject = fields[0];
  const startDate = fields[1];
  const startTime = fields[2];
  const endDate = fields[3];
  const endTime = fields[4];
  const location = fields[5];
  // Unescape literal \n back to actual newlines
  const description = fields[6]?.replace(/\\n/g, '\n');
  const organizer = fields[7];
  const isAllDay = fields[8] === 'Yes';
  const importance = fields[9] as 'low' | 'normal' | 'high';
  const busyStatus = fields[10] as 'free' | 'tentative' | 'busy' | 'out-of-office';
  const sensitivity = fields[11] as 'normal' | 'personal' | 'private' | 'confidential';
  const isRecurring = fields[12] === 'Yes';
  let recurrencePattern = fields[13];
  const reminder = fields[14] ? parseInt(fields[14]) : undefined;

  // Fix: Yearly recurrence with INTERVAL=12 is invalid - should be INTERVAL=1 or omitted
  // Outlook stores yearly as 12 months, but iCalendar expects 1 year
  if (recurrencePattern && recurrencePattern.includes('FREQ=YEARLY') && recurrencePattern.includes('INTERVAL=12')) {
    recurrencePattern = recurrencePattern.replace(/INTERVAL=12;?/, '');
    // Clean up any double semicolons
    recurrencePattern = recurrencePattern.replace(/;;/g, ';');
  }

  // Fix: UNTIL must include time component to match DTSTART format (RFC 5545)
  // Convert UNTIL=YYYYMMDD to UNTIL=YYYYMMDDTHHMMSSZ
  if (recurrencePattern && recurrencePattern.includes('UNTIL=')) {
    recurrencePattern = recurrencePattern.replace(/UNTIL=(\d{8})(?!T)/g, 'UNTIL=$1T235959Z');
  }

  // Fix: Daily recurrence intervals stored as minutes (MS-OXOCAL spec)
  // Convert INTERVAL=1440 (minutes) to INTERVAL=1 (day), INTERVAL=2880 to INTERVAL=2, etc.
  if (recurrencePattern && recurrencePattern.includes('FREQ=DAILY') && recurrencePattern.includes('INTERVAL=')) {
    recurrencePattern = recurrencePattern.replace(/INTERVAL=(\d+)/g, (match, minutes) => {
      const days = Math.floor(parseInt(minutes) / 1440);
      return days > 1 ? `INTERVAL=${days}` : ''; // Omit INTERVAL=1 (it's the default)
    });
    // Clean up any leftover semicolons
    recurrencePattern = recurrencePattern.replace(/;;+/g, ';').replace(/;$/, '');
  }

  // Fix: Yearly recurrence with BYMONTHDAY needs BYMONTH (RFC 5545 compliance)
  // Google Calendar requires explicit BYMONTH when using BYMONTHDAY with FREQ=YEARLY
  if (recurrencePattern && recurrencePattern.includes('FREQ=YEARLY') && recurrencePattern.includes('BYMONTHDAY=')) {
    // Check if BYMONTH is already present
    if (!recurrencePattern.includes('BYMONTH=')) {
      // Extract month from startDate (format: YYYY-MM-DD)
      const month = parseInt(startDate.substring(5, 7)); // Extract MM from YYYY-MM-DD
      // Insert BYMONTH before BYMONTHDAY
      recurrencePattern = recurrencePattern.replace(/BYMONTHDAY=/, `BYMONTH=${month};BYMONTHDAY=`);
    }
  }

  // Fix: Yearly recurrence with BYDAY needs BYMONTH (RFC 5545 compliance)
  // Without BYMONTH, "FREQ=YEARLY;BYDAY=-1SA" means "last Saturday of the year" (December)
  // With BYMONTH=1, it means "last Saturday of January"
  // This fixes legacy data exported before this fix was added to calendar-extractor.ts
  if (
    recurrencePattern &&
    recurrencePattern.includes('FREQ=YEARLY') &&
    recurrencePattern.includes('BYDAY=') &&
    !recurrencePattern.includes('BYMONTH=')
  ) {
    // Extract month from startDate (format: YYYY-MM-DD)
    const month = parseInt(startDate.substring(5, 7)); // Extract MM from YYYY-MM-DD
    // Insert BYMONTH after FREQ=YEARLY
    recurrencePattern = recurrencePattern.replace(/FREQ=YEARLY;?/, `FREQ=YEARLY;BYMONTH=${month};`);
  }

  // Fix: Suspicious UNTIL=2100 patterns (corrupted Outlook data)
  // Cap far-future UNTIL dates to reasonable spans based on frequency
  if (recurrencePattern && recurrencePattern.includes('UNTIL=2100')) {
    const startDateObj = new Date(startDate);
    const yearsSpan = 2100 - startDateObj.getFullYear();

    // Daily recurrence: cap at 5 years
    if (recurrencePattern.includes('FREQ=DAILY') && yearsSpan > 5) {
      const cappedYear = startDateObj.getFullYear() + 5;
      const cappedDate = recurrencePattern.match(/UNTIL=\d{4}(\d{4}T\d{6}Z)/);
      const dateSuffix = cappedDate ? cappedDate[1] : '1231T235959Z';
      recurrencePattern = recurrencePattern.replace(/UNTIL=2100\d{4}T\d{6}Z/, `UNTIL=${cappedYear}${dateSuffix}`);
      console.warn(`⚠️  Capped daily recurrence for "${subject}" to ${cappedYear} (was 2100)`);
    }

    // Weekly recurrence: cap at 10 years
    if (recurrencePattern.includes('FREQ=WEEKLY') && yearsSpan > 10) {
      const cappedYear = startDateObj.getFullYear() + 10;
      const cappedDate = recurrencePattern.match(/UNTIL=\d{4}(\d{4}T\d{6}Z)/);
      const dateSuffix = cappedDate ? cappedDate[1] : '1231T235959Z';
      recurrencePattern = recurrencePattern.replace(/UNTIL=2100\d{4}T\d{6}Z/, `UNTIL=${cappedYear}${dateSuffix}`);
      console.warn(`⚠️  Capped weekly recurrence for "${subject}" to ${cappedYear} (was 2100)`);
    }

    // Monthly recurrence: cap at 20 years
    if (recurrencePattern.includes('FREQ=MONTHLY') && yearsSpan > 20) {
      const cappedYear = startDateObj.getFullYear() + 20;
      const cappedDate = recurrencePattern.match(/UNTIL=\d{4}(\d{4}T\d{6}Z)/);
      const dateSuffix = cappedDate ? cappedDate[1] : '1231T235959Z';
      recurrencePattern = recurrencePattern.replace(/UNTIL=2100\d{4}T\d{6}Z/, `UNTIL=${cappedYear}${dateSuffix}`);
      console.warn(`⚠️  Capped monthly recurrence for "${subject}" to ${cappedYear} (was 2100)`);
    }

    // Yearly recurrence: allow up to 100 years (birthdays/anniversaries are legitimate)
  }

  // Parse dates
  let startDateTime: Date;
  let endDateTime: Date;

  if (isAllDay || startTime === 'All Day') {
    // All-day event - parse just the date
    startDateTime = new Date(startDate);
    endDateTime = new Date(endDate);
  } else {
    // Timed event - combine date and time
    startDateTime = new Date(`${startDate}T${startTime}Z`);
    endDateTime = new Date(`${endDate}T${endTime}Z`);
  }

  return {
    subject,
    startTime: startDateTime,
    endTime: endDateTime,
    location: location || undefined,
    description: description || undefined,
    organizer: organizer || undefined,
    isAllDay,
    isRecurring,
    recurrencePattern: recurrencePattern || undefined,
    reminder,
    importance: importance || undefined,
    busyStatus: busyStatus || undefined,
    sensitivity: sensitivity || undefined,
  };
}

/**
 * Split large calendar into multiple ICS files for Google Calendar compatibility
 * Google Calendar has a 499-event import limit per file
 */
async function exportToICS() {
  const csvPath = 'output/calendar-export.csv';

  // Check if CSV file exists
  if (!fs.existsSync(csvPath)) {
    console.error('CSV file not found. Please run export-to-csv.ts first.');
    console.error(`Expected location: ${csvPath}`);
    process.exit(1);
  }

  console.log('Reading CSV file...');
  const csvContent = fs.readFileSync(csvPath, 'utf-8');
  const lines = csvContent.split('\n').filter((line) => line.trim());

  if (lines.length <= 1) {
    console.log('No entries to export (CSV is empty or has only headers)');
    return;
  }

  console.log(`Found ${lines.length - 1} entries`);

  try {
    // Skip header line and parse entries
    console.log('Converting entries...');
    const entries: CalendarEntry[] = [];

    for (let i = 1; i < lines.length; i++) {
      const fields = parseCSVLine(lines[i]);
      if (fields.length >= 15) {
        // Ensure we have all expected fields
        const entry = rowToEntry(fields);
        entries.push(entry);
      }
    }

    console.log(`Converted ${entries.length} entries`);

    // Automatically split if needed (default: 499 events per file)
    const converter = new ICalConverter();
    await converter.convertAndSaveSplit(entries, 'output/calendar-export.ics', {
      calendarName: 'Exported Calendar',
      timezone: 'UTC',
    });
  } catch (error) {
    console.error('Export failed:', (error as Error).message);
    throw error;
  }
}

exportToICS().catch((error) => {
  console.error('Error:', error.message);
  process.exit(1);
});
