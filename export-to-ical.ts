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

async function exportToICS() {
  const csvPath = 'calendar-export.csv';
  const outputPath = 'calendar-export.ics';

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

    // Use ICalConverter to generate ICS file
    console.log('Generating iCalendar file...');
    const converter = new ICalConverter();
    const calendar = converter.convert(entries, {
      calendarName: 'Exported Calendar',
      timezone: 'UTC',
    });

    // Save to file
    await converter.saveToFile(calendar, outputPath);

    console.log(`\n✓ Successfully exported ${entries.length} entries to ${outputPath}`);
    console.log(`\nYou can import this file into any calendar application that supports iCalendar format.`);
  } catch (error) {
    console.error('Export failed:', (error as Error).message);
    throw error;
  }
}

exportToICS().catch((error) => {
  console.error('Error:', error.message);
  process.exit(1);
});
