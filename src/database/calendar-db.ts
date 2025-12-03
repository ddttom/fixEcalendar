import Database from 'better-sqlite3';
import * as path from 'path';
import * as fs from 'fs';
import * as crypto from 'crypto';
import { CalendarEntry } from '../parser/types';
import { logger } from '../utils/logger';

export interface DBCalendarEntry extends CalendarEntry {
  id?: number;
  hash: string;
  sourceFile: string;
  processedAt: Date;
}

export class CalendarDatabase {
  private db: Database.Database;
  private dbPath: string;

  constructor(dbPath?: string) {
    // Default to .fixecalendar.db in current directory
    this.dbPath = dbPath || path.join(process.cwd(), '.fixecalendar.db');

    logger.debug(`Initializing database at: ${this.dbPath}`);

    // Ensure directory exists
    const dir = path.dirname(this.dbPath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    this.db = new Database(this.dbPath);
    this.db.pragma('journal_mode = WAL'); // Better concurrency
    this.db.pragma('synchronous = NORMAL'); // Better performance

    this.initializeSchema();
    logger.info(`Database initialized: ${this.dbPath}`);
  }

  private initializeSchema(): void {
    // Create main calendar entries table
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS calendar_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hash TEXT UNIQUE NOT NULL,
        uid TEXT,
        subject TEXT NOT NULL,
        start_time INTEGER NOT NULL,
        end_time INTEGER NOT NULL,
        location TEXT,
        description TEXT,
        organizer TEXT,
        is_all_day INTEGER DEFAULT 0,
        is_recurring INTEGER DEFAULT 0,
        recurrence_pattern TEXT,
        reminder INTEGER,
        importance TEXT,
        busy_status TEXT,
        sensitivity TEXT,
        source_file TEXT NOT NULL,
        processed_at INTEGER NOT NULL,
        created_at INTEGER DEFAULT (strftime('%s', 'now'))
      );

      CREATE INDEX IF NOT EXISTS idx_hash ON calendar_entries(hash);
      CREATE INDEX IF NOT EXISTS idx_uid ON calendar_entries(uid);
      CREATE INDEX IF NOT EXISTS idx_start_time ON calendar_entries(start_time);
      CREATE INDEX IF NOT EXISTS idx_source_file ON calendar_entries(source_file);
      CREATE INDEX IF NOT EXISTS idx_subject_time ON calendar_entries(subject, start_time, end_time);

      CREATE TABLE IF NOT EXISTS attendees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_id INTEGER NOT NULL,
        name TEXT,
        email TEXT,
        type TEXT,
        FOREIGN KEY(entry_id) REFERENCES calendar_entries(id) ON DELETE CASCADE
      );

      CREATE INDEX IF NOT EXISTS idx_entry_id ON attendees(entry_id);

      CREATE TABLE IF NOT EXISTS processing_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_file TEXT NOT NULL,
        entries_found INTEGER DEFAULT 0,
        entries_added INTEGER DEFAULT 0,
        entries_skipped INTEGER DEFAULT 0,
        started_at INTEGER NOT NULL,
        completed_at INTEGER,
        status TEXT DEFAULT 'processing',
        error TEXT
      );

      CREATE INDEX IF NOT EXISTS idx_processing_status ON processing_log(status);
    `);
  }

  /**
   * Generate a hash for an entry to detect duplicates
   */
  private generateEntryHash(entry: CalendarEntry): string {
    // Create a unique hash based on key fields
    const data = [
      entry.uid || '',
      entry.subject,
      entry.startTime.toISOString(),
      entry.endTime.toISOString(),
      entry.location || '',
      entry.organizer || '',
    ].join('|');

    return crypto.createHash('sha256').update(data).digest('hex');
  }

  /**
   * Check if an entry already exists in the database
   */
  isDuplicate(entry: CalendarEntry): boolean {
    const hash = this.generateEntryHash(entry);

    const stmt = this.db.prepare('SELECT COUNT(*) as count FROM calendar_entries WHERE hash = ?');
    const result = stmt.get(hash) as { count: number };

    return result.count > 0;
  }

  /**
   * Add a calendar entry to the database (with deduplication)
   */
  addEntry(entry: CalendarEntry, sourceFile: string): boolean {
    const hash = this.generateEntryHash(entry);

    // Check for duplicate
    if (this.isDuplicate(entry)) {
      logger.debug(`Duplicate entry skipped: ${entry.subject} (${entry.startTime.toISOString()})`);
      return false;
    }

    const insertEntry = this.db.prepare(`
      INSERT INTO calendar_entries (
        hash, uid, subject, start_time, end_time, location, description,
        organizer, is_all_day, is_recurring, recurrence_pattern, reminder,
        importance, busy_status, sensitivity, source_file, processed_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    const insertAttendee = this.db.prepare(`
      INSERT INTO attendees (entry_id, name, email, type) VALUES (?, ?, ?, ?)
    `);

    const transaction = this.db.transaction(() => {
      // Insert main entry
      const info = insertEntry.run(
        hash,
        entry.uid || null,
        entry.subject,
        entry.startTime.getTime(),
        entry.endTime.getTime(),
        entry.location || null,
        entry.description || null,
        entry.organizer || null,
        entry.isAllDay ? 1 : 0,
        entry.isRecurring ? 1 : 0,
        entry.recurrencePattern || null,
        entry.reminder || null,
        entry.importance || null,
        entry.busyStatus || null,
        entry.sensitivity || null,
        sourceFile,
        Date.now()
      );

      // Insert attendees if present
      if (entry.attendees && entry.attendees.length > 0) {
        const entryId = info.lastInsertRowid;
        for (const attendee of entry.attendees) {
          insertAttendee.run(
            entryId,
            attendee.name || null,
            attendee.email || null,
            attendee.type || null
          );
        }
      }
    });

    try {
      transaction();
      return true;
    } catch (error) {
      logger.error(`Failed to add entry: ${entry.subject}`, error as Error);
      return false;
    }
  }

  /**
   * Add multiple entries in a single transaction (more efficient)
   */
  addEntries(entries: CalendarEntry[], sourceFile: string): { added: number; skipped: number } {
    let added = 0;
    let skipped = 0;

    const transaction = this.db.transaction(() => {
      for (const entry of entries) {
        if (this.addEntry(entry, sourceFile)) {
          added++;
        } else {
          skipped++;
        }
      }
    });

    transaction();

    return { added, skipped };
  }

  /**
   * Get all entries from the database
   */
  getAllEntries(): CalendarEntry[] {
    const stmt = this.db.prepare(`
      SELECT * FROM calendar_entries ORDER BY start_time ASC
    `);

    const rows = stmt.all() as any[];
    return rows.map(row => this.rowToEntry(row));
  }

  /**
   * Get entries within a date range
   */
  getEntriesInRange(startDate?: Date, endDate?: Date): CalendarEntry[] {
    let query = 'SELECT * FROM calendar_entries WHERE 1=1';
    const params: number[] = [];

    if (startDate) {
      query += ' AND start_time >= ?';
      params.push(startDate.getTime());
    }

    if (endDate) {
      query += ' AND start_time <= ?';
      params.push(endDate.getTime());
    }

    query += ' ORDER BY start_time ASC';

    const stmt = this.db.prepare(query);
    const rows = stmt.all(...params) as any[];

    return rows.map(row => this.rowToEntry(row));
  }

  /**
   * Get entries from a specific source file
   */
  getEntriesBySource(sourceFile: string): CalendarEntry[] {
    const stmt = this.db.prepare(`
      SELECT * FROM calendar_entries WHERE source_file = ? ORDER BY start_time ASC
    `);

    const rows = stmt.all(sourceFile) as any[];
    return rows.map(row => this.rowToEntry(row));
  }

  /**
   * Get database statistics
   */
  getStats(): {
    totalEntries: number;
    sourceFiles: number;
    dateRange: { earliest?: Date; latest?: Date };
    entriesBySource: Map<string, number>;
  } {
    const totalStmt = this.db.prepare('SELECT COUNT(*) as count FROM calendar_entries');
    const total = (totalStmt.get() as { count: number }).count;

    const sourceStmt = this.db.prepare('SELECT COUNT(DISTINCT source_file) as count FROM calendar_entries');
    const sources = (sourceStmt.get() as { count: number }).count;

    const rangeStmt = this.db.prepare(`
      SELECT MIN(start_time) as earliest, MAX(start_time) as latest
      FROM calendar_entries
    `);
    const range = rangeStmt.get() as { earliest?: number; latest?: number };

    const bySourceStmt = this.db.prepare(`
      SELECT source_file, COUNT(*) as count
      FROM calendar_entries
      GROUP BY source_file
    `);
    const bySource = bySourceStmt.all() as Array<{ source_file: string; count: number }>;

    const entriesBySource = new Map<string, number>();
    bySource.forEach(row => {
      entriesBySource.set(row.source_file, row.count);
    });

    return {
      totalEntries: total,
      sourceFiles: sources,
      dateRange: {
        earliest: range.earliest ? new Date(range.earliest) : undefined,
        latest: range.latest ? new Date(range.latest) : undefined,
      },
      entriesBySource,
    };
  }

  /**
   * Log processing of a file
   */
  logProcessing(
    sourceFile: string,
    entriesFound: number,
    entriesAdded: number,
    entriesSkipped: number
  ): void {
    const stmt = this.db.prepare(`
      INSERT INTO processing_log (
        source_file, entries_found, entries_added, entries_skipped,
        started_at, completed_at, status
      ) VALUES (?, ?, ?, ?, ?, ?, 'completed')
    `);

    stmt.run(
      sourceFile,
      entriesFound,
      entriesAdded,
      entriesSkipped,
      Date.now(),
      Date.now()
    );
  }

  /**
   * Clear all entries (with confirmation)
   */
  clear(): void {
    this.db.exec('DELETE FROM attendees');
    this.db.exec('DELETE FROM calendar_entries');
    this.db.exec('DELETE FROM processing_log');
    logger.info('Database cleared');
  }

  /**
   * Convert database row to CalendarEntry
   */
  private rowToEntry(row: any): CalendarEntry {
    // Get attendees for this entry
    const attendeeStmt = this.db.prepare('SELECT * FROM attendees WHERE entry_id = ?');
    const attendeeRows = attendeeStmt.all(row.id) as any[];

    return {
      uid: row.uid,
      subject: row.subject,
      startTime: new Date(row.start_time),
      endTime: new Date(row.end_time),
      location: row.location || undefined,
      description: row.description || undefined,
      organizer: row.organizer || undefined,
      attendees: attendeeRows.length > 0
        ? attendeeRows.map(a => ({
            name: a.name || undefined,
            email: a.email || undefined,
            type: a.type || undefined,
          }))
        : undefined,
      isAllDay: row.is_all_day === 1,
      isRecurring: row.is_recurring === 1,
      recurrencePattern: row.recurrence_pattern || undefined,
      reminder: row.reminder || undefined,
      importance: row.importance || undefined,
      busyStatus: row.busy_status || undefined,
      sensitivity: row.sensitivity || undefined,
    };
  }

  /**
   * Close the database connection
   */
  close(): void {
    this.db.close();
    logger.debug('Database connection closed');
  }

  /**
   * Get database file path
   */
  getPath(): string {
    return this.dbPath;
  }

  /**
   * Optimize database (vacuum and analyze)
   */
  optimize(): void {
    logger.info('Optimizing database...');
    this.db.exec('VACUUM');
    this.db.exec('ANALYZE');
    logger.info('Database optimized');
  }
}
