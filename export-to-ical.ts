#!/usr/bin/env ts-node

/**
 * Export calendar entries from CSV to iCalendar (ICS) format
 */

import * as fs from 'fs';
import { CalendarEntry } from './src/parser/types';
import { ICalConverter } from './src/converter/ical-converter';

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
async function splitAndSaveCalendar(
  entries: CalendarEntry[],
  eventsPerFile: number = 499
): Promise<void> {
  const totalEvents = entries.length;
  const numFiles = Math.ceil(totalEvents / eventsPerFile);

  console.log(`\nGenerating ${numFiles} iCalendar file${numFiles > 1 ? 's' : ''} for Google Calendar...`);
  console.log(`Total events: ${totalEvents} (${eventsPerFile} events per file)\n`);

  for (let fileNum = 0; fileNum < numFiles; fileNum++) {
    const startIdx = fileNum * eventsPerFile;
    const endIdx = Math.min((fileNum + 1) * eventsPerFile, totalEvents);
    const chunkEntries = entries.slice(startIdx, endIdx);

    const outputPath = numFiles === 1
      ? 'calendar-export.ics'
      : `calendar-part-${fileNum + 1}-of-${numFiles}.ics`;

    const converter = new ICalConverter();
    const calendarName = numFiles === 1
      ? 'Exported Calendar'
      : `Exported Calendar (Part ${fileNum + 1} of ${numFiles})`;

    const calendar = converter.convert(chunkEntries, {
      calendarName,
      timezone: 'UTC',
    });

    await converter.saveToFile(calendar, outputPath);
    console.log(`✓ Created ${outputPath} with ${chunkEntries.length} events`);
  }

  console.log(`\n✓ Successfully created ${numFiles} file${numFiles > 1 ? 's' : ''}`);

  if (numFiles > 1) {
    console.log(`\n📋 Import Instructions:`);
    console.log(`   Import split files sequentially into Google Calendar:`);
    console.log(`     1. Import calendar-part-1-of-${numFiles}.ics`);
    console.log(`     2. Wait for import to complete`);
    console.log(`     3. Import calendar-part-2-of-${numFiles}.ics`);
    console.log(`     4. Repeat for all ${numFiles} files`);
  } else {
    console.log(`\nYou can import calendar-export.ics into Google Calendar or any RFC 5545-compliant calendar application.`);
  }
}

async function exportToICS() {
  const csvPath = 'calendar-export.csv';

  // Check if CSV file exists
  if (!fs.existsSync(csvPath)) {
    console.error('CSV file not found. Please run export-to-csv.ts first.');
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
    await splitAndSaveCalendar(entries);
  } catch (error) {
    console.error('Export failed:', (error as Error).message);
    throw error;
  }
}

exportToICS().catch((error) => {
  console.error('Error:', error.message);
  process.exit(1);
});
