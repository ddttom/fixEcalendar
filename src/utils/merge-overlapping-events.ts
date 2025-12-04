#!/usr/bin/env ts-node
/**
 * Merge Overlapping Events Script
 *
 * Issue: Some events have multiple entries on the same day with overlapping or adjacent times
 * Examples:
 *   - Aquafit 09:30 (08:00-10:00) + Aquafit 11:05 (09:30-11:30) → Aquafit (08:00-11:30)
 *   - Board Meeting (09:30-12:00) + Board Meeting (14:00-17:00) → Board Meeting (09:30-17:00)
 *
 * Strategy:
 * 1. Find events with same subject on same day with different start times
 * 2. Normalize subject (remove time prefixes like "09:30", "11:00")
 * 3. Merge into single event with earliest start and latest end time
 * 4. Keep the entry with the longest description
 * 5. Delete duplicate entries
 *
 * Usage: npx ts-node src/utils/merge-overlapping-events.ts [--dry-run]
 */

import Database from 'better-sqlite3';
import { join } from 'path';

interface OverlappingEvent {
  event_date: string;
  subject: string;
  event_count: number;
  distinct_times: number;
  earliest_start: number;
  latest_end: number;
}

interface EventDetail {
  id: number;
  subject: string;
  start_time: number;
  end_time: number;
  description: string | null;
  location: string | null;
  organizer: string | null;
  recurrence_pattern: string | null;
}

function normalizeSubject(subject: string): string {
  // Remove time prefixes like "09:30", "11:00", "14:00 - 15:00"
  let normalized = subject
    .replace(/^\d{1,2}:\d{2}\s+/, '')           // Remove "09:30 "
    .replace(/^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}\s+/, '') // Remove "14:00 - 15:00 "
    .replace(/\s+\d{1,2}:\d{2}$/, '')           // Remove trailing " 09:30"
    .trim();

  // Remove trailing punctuation and whitespace
  normalized = normalized.replace(/[.,:;!?\s]+$/, '').trim();

  return normalized;
}

function areSubjectsSimilar(subject1: string, subject2: string): boolean {
  const norm1 = normalizeSubject(subject1.toLowerCase());
  const norm2 = normalizeSubject(subject2.toLowerCase());

  // Exact match after normalization
  if (norm1 === norm2) return true;

  // Check if one contains the other (handle variations)
  if (norm1.includes(norm2) || norm2.includes(norm1)) return true;

  return false;
}

function doEventsOverlap(start1: number, end1: number, start2: number, end2: number): boolean {
  // Events overlap if one starts before the other ends
  return (start1 < end2 && start2 < end1);
}

function selectBestDescription(events: EventDetail[]): string {
  // Select the longest non-junk description
  const descriptions = events
    .map(e => e.description || '')
    .filter(d => {
      const trimmed = d.trim();
      // Filter out HTML/CSS junk
      if (!trimmed) return false;
      if (trimmed.startsWith('false\n')) return false;
      if (trimmed.includes('/* Style Definitions */')) return false;
      if (trimmed.includes('table.MsoNormal')) return false;
      return true;
    });

  // Return the longest description, or empty string
  return descriptions.sort((a, b) => b.length - a.length)[0] || '';
}

async function mergeOverlappingEvents(dryRun: boolean = false) {
  const dbPath = join(process.cwd(), '.fixecalendar.db');
  console.log(`\nOpening database: ${dbPath}`);
  console.log(`Mode: ${dryRun ? 'DRY RUN (no changes)' : 'LIVE (will modify database)'}\n`);

  const db = new Database(dbPath);

  try {
    // First, get all events grouped by date
    // We'll process client-side to handle subject normalization
    const allEventsQuery = `
      SELECT
        id,
        subject,
        DATE(start_time/1000, 'unixepoch') as event_date,
        start_time,
        end_time,
        description,
        location,
        organizer,
        recurrence_pattern
      FROM calendar_entries
      WHERE is_all_day = 0  -- Only timed events
      ORDER BY event_date, start_time
    `;

    const allEvents = db.prepare(allEventsQuery).all() as (EventDetail & { event_date: string })[];

    // Group by date and normalized subject
    const groupedByDate = new Map<string, Map<string, (EventDetail & { event_date: string })[]>>();

    for (const event of allEvents) {
      const date = event.event_date;
      const normalized = normalizeSubject(event.subject);

      if (!groupedByDate.has(date)) {
        groupedByDate.set(date, new Map());
      }

      const dateGroup = groupedByDate.get(date)!;
      if (!dateGroup.has(normalized)) {
        dateGroup.set(normalized, []);
      }

      dateGroup.get(normalized)!.push(event);
    }

    // Find candidates with multiple events
    const candidates: Array<{ date: string; normalized: string; events: (EventDetail & { event_date: string })[] }> = [];

    for (const [date, subjectGroups] of groupedByDate.entries()) {
      for (const [normalized, events] of subjectGroups.entries()) {
        if (events.length > 1) {
          candidates.push({ date, normalized, events });
        }
      }
    }

    console.log(`Found ${candidates.length} days with potential overlapping events\n`);

    if (candidates.length === 0) {
      console.log('✓ No overlapping events found!\n');
      return;
    }

    let mergedCount = 0;
    let deletedCount = 0;

    for (const candidate of candidates) {
      const events = candidate.events;
      const eventDate = candidate.date;
      const normalizedSubject = candidate.normalized;

      // Check if events overlap or are adjacent
      let hasOverlap = false;
      for (let i = 0; i < events.length - 1; i++) {
        if (doEventsOverlap(events[i].start_time, events[i].end_time,
                           events[i + 1].start_time, events[i + 1].end_time)) {
          hasOverlap = true;
          break;
        }
      }

      if (!hasOverlap) {
        console.log(`⏭️  Skipping "${normalizedSubject}" on ${eventDate} - events don't overlap`);
        continue;
      }

      // Merge the events
      const mergedSubject = normalizedSubject;
      const mergedStart = Math.min(...events.map(e => e.start_time));
      const mergedEnd = Math.max(...events.map(e => e.end_time));
      const mergedDescription = selectBestDescription(events);
      const mergedLocation = events.find(e => e.location)?.location || null;
      const mergedOrganizer = events.find(e => e.organizer)?.organizer || null;

      // Keep the first event ID, delete the rest
      const keepId = events[0].id;
      const deleteIds = events.slice(1).map(e => e.id);

      console.log(`📅 ${eventDate} - "${normalizedSubject}"`);
      console.log(`   Events found: ${events.length}`);
      events.forEach(e => {
        const start = new Date(e.start_time).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
        const end = new Date(e.end_time).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
        console.log(`     - "${e.subject}" (${start} - ${end})`);
      });

      const mergedStartTime = new Date(mergedStart).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
      const mergedEndTime = new Date(mergedEnd).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
      console.log(`   Merged: "${mergedSubject}" (${mergedStartTime} - ${mergedEndTime})`);

      if (!dryRun) {
        // Update the first event
        db.prepare(`
          UPDATE calendar_entries
          SET subject = ?,
              start_time = ?,
              end_time = ?,
              description = ?,
              location = COALESCE(?, location),
              organizer = COALESCE(?, organizer)
          WHERE id = ?
        `).run(mergedSubject, mergedStart, mergedEnd, mergedDescription, mergedLocation, mergedOrganizer, keepId);

        // Delete the duplicate events
        if (deleteIds.length > 0) {
          const placeholders = deleteIds.map(() => '?').join(',');
          db.prepare(`DELETE FROM calendar_entries WHERE id IN (${placeholders})`).run(...deleteIds);
          deletedCount += deleteIds.length;
        }

        mergedCount++;
        console.log(`   ✓ Merged and deleted ${deleteIds.length} duplicate(s)\n`);
      } else {
        console.log(`   [DRY RUN] Would merge and delete ${deleteIds.length} duplicate(s)\n`);
        mergedCount++;
      }
    }

    console.log('==================');
    if (dryRun) {
      console.log(`[DRY RUN] Would merge ${mergedCount} groups and delete ${deletedCount} duplicate entries`);
      console.log('\nRun without --dry-run to apply changes\n');
    } else {
      console.log(`✓ Successfully merged ${mergedCount} groups and deleted ${deletedCount} duplicate entries\n`);
      console.log('Next steps:');
      console.log('1. Run: npx ts-node export-to-csv.ts');
      console.log('2. Run: npx ts-node export-to-ical.ts\n');
    }
  } catch (error) {
    console.error('Error during merge:', error);
    throw error;
  } finally {
    db.close();
  }
}

// Export the function for use by other scripts
export { mergeOverlappingEvents };

// CLI entry point
if (require.main === module) {
  // Parse command line arguments
  const dryRun = process.argv.includes('--dry-run');

  // Run the merge
  mergeOverlappingEvents(dryRun).catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}
