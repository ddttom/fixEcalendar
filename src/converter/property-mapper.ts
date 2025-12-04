import type { ICalEventData, ICalAttendeeData } from 'ical-generator';
import {
  ICalEventStatus,
  ICalEventBusyStatus,
  ICalAttendeeRole,
  ICalAlarmType,
  ICalEventClass,
  ICalEventRepeatingFreq,
} from 'ical-generator';
import { CalendarEntry } from '../parser/types';
import { logger } from '../utils/logger';

export class PropertyMapper {
  private standardizeBirthdaySubject(
    subject: string,
    startDate: Date,
    recurrencePattern: string | undefined
  ): string {
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

  mapToICalEvent(entry: CalendarEntry): ICalEventData {
    let startTime = entry.startTime;
    let endTime = entry.endTime;
    let isAllDay = entry.isAllDay;

    // Special handling for birthdays/anniversaries - force them to be all-day events with correct date
    const subject = (entry.subject || '').toLowerCase();
    const isBirthdayOrAnniversary = subject.includes('birthday') || subject.includes('anniversary');

    if (isBirthdayOrAnniversary && entry.recurrencePattern) {
      // Extract day from BYMONTHDAY in recurrence pattern
      const byMonthDayMatch = entry.recurrencePattern.match(/BYMONTHDAY=(\d+)/);
      if (byMonthDayMatch) {
        const correctDay = parseInt(byMonthDayMatch[1]);
        // Create date with the correct day from recurrence pattern
        const baseDate = new Date(entry.startTime);
        const year = baseDate.getFullYear();
        const month = baseDate.getMonth();

        startTime = new Date(year, month, correctDay);
        endTime = new Date(year, month, correctDay); // Same day for all-day events
        isAllDay = true; // Force as all-day event
      }
    }

    // Handle all-day events - normalize to UTC midnight for proper DATE formatting
    if (isAllDay) {
      // Convert to UTC date at midnight (remove time component)
      const startDate = new Date(startTime);
      const endDate = new Date(endTime);

      // Create new dates at UTC midnight using Date.UTC
      // This ensures VALUE=DATE formatting shows the correct date
      startTime = new Date(
        Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate())
      );
      endTime = new Date(Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()));
    }

    // Standardize birthday/anniversary subject format
    const standardizedSubject = this.standardizeBirthdaySubject(
      entry.subject,
      new Date(entry.startTime),
      entry.recurrencePattern
    );

    const eventData: ICalEventData = {
      start: startTime,
      end: endTime,
      summary: standardizedSubject,
      description: entry.description,
      location: entry.location,
      id: entry.uid,
    };

    // Handle all-day events
    if (isAllDay) {
      eventData.allDay = true;
    }

    // Map organizer
    if (entry.organizer) {
      const org = this.mapOrganizer(entry.organizer);
      if (org.email) {
        eventData.organizer = { name: org.name || '', email: org.email };
      }
    }

    // Map attendees
    if (entry.attendees && entry.attendees.length > 0) {
      eventData.attendees = entry.attendees.map((a) => this.mapAttendee(a)).filter((a) => a.email);
    }

    // Map status
    if (entry.busyStatus) {
      eventData.status = this.mapStatus(entry.busyStatus);
      eventData.busystatus = this.mapBusyStatus(entry.busyStatus);
    }

    // Map priority
    if (entry.importance) {
      eventData.priority = this.mapPriority(entry.importance);
    }

    // Map reminder/alarm
    if (entry.reminder !== undefined) {
      eventData.alarms = [this.mapReminder(entry.reminder)];
    }

    // Map recurrence pattern to RRULE
    if (entry.isRecurring && entry.recurrencePattern) {
      // The recurrence pattern should already be in RRULE format from the extractor
      const rrule = entry.recurrencePattern;

      // Check if it starts with 'RRULE' or 'RRULE:'
      if (rrule.startsWith('RRULE:')) {
        eventData.repeating = rrule.substring(6); // Remove 'RRULE:' prefix
      } else if (rrule.startsWith('RRULE;')) {
        eventData.repeating = rrule.substring(6); // Remove 'RRULE;' prefix
      } else if (rrule === 'RECURRING') {
        // Fallback for unparsed recurrences - just mark as recurring daily
        logger.debug(`Note: Recurring event "${entry.subject}" has unparseable recurrence pattern`);
        eventData.repeating = { freq: ICalEventRepeatingFreq.DAILY };
      } else {
        // Parse RRULE string into object format expected by ical-generator
        eventData.repeating = this.parseRRuleString(rrule);
      }
    }

    // Map sensitivity/class
    if (entry.sensitivity) {
      eventData.class = this.mapClass(entry.sensitivity);
    }

    return eventData;
  }

  private mapOrganizer(organizer: string): { name?: string; email?: string } {
    // Try to parse "Name <email>" format
    const match = organizer.match(/^(.+?)\s*<(.+?)>$/);

    if (match) {
      return {
        name: match[1].trim(),
        email: match[2].trim(),
      };
    }

    // Check if it's just an email
    if (organizer.includes('@')) {
      return { email: organizer };
    }

    // Otherwise, treat as name
    return { name: organizer };
  }

  private mapAttendee(attendee: {
    name?: string;
    email?: string;
    type?: 'required' | 'optional' | 'resource';
  }): ICalAttendeeData {
    const attendeeData: ICalAttendeeData = {
      email: attendee.email || '',
      name: attendee.name,
      rsvp: false,
    };

    // Map attendee role
    switch (attendee.type) {
      case 'required':
        attendeeData.role = ICalAttendeeRole.REQ;
        break;
      case 'optional':
        attendeeData.role = ICalAttendeeRole.OPT;
        break;
      case 'resource':
        attendeeData.role = ICalAttendeeRole.NON;
        break;
    }

    return attendeeData;
  }

  private mapStatus(busyStatus: string): ICalEventStatus {
    // Map Outlook busy status to iCal STATUS property
    switch (busyStatus) {
      case 'tentative':
        return ICalEventStatus.TENTATIVE;
      case 'free':
      case 'busy':
      case 'out-of-office':
      default:
        return ICalEventStatus.CONFIRMED;
    }
  }

  private mapBusyStatus(busyStatus: string): ICalEventBusyStatus {
    // Map to iCal TRANSP/FREEBUSY property
    switch (busyStatus) {
      case 'free':
        return ICalEventBusyStatus.FREE;
      case 'tentative':
        return ICalEventBusyStatus.TENTATIVE;
      case 'out-of-office':
        return ICalEventBusyStatus.OOF;
      case 'busy':
      default:
        return ICalEventBusyStatus.BUSY;
    }
  }

  private mapPriority(importance: 'low' | 'normal' | 'high'): number {
    // iCal priority: 1 = highest, 5 = medium, 9 = lowest
    switch (importance) {
      case 'high':
        return 1;
      case 'low':
        return 9;
      case 'normal':
      default:
        return 5;
    }
  }

  private mapReminder(minutesBefore: number) {
    // Create a VALARM component
    return {
      type: ICalAlarmType.display,
      trigger: minutesBefore * 60, // Convert minutes to seconds
      description: 'Reminder',
    };
  }

  private mapClass(sensitivity: string): ICalEventClass {
    // Map sensitivity to iCal CLASS property
    switch (sensitivity) {
      case 'private':
        return ICalEventClass.PRIVATE;
      case 'confidential':
        return ICalEventClass.CONFIDENTIAL;
      case 'personal':
      case 'normal':
      default:
        return ICalEventClass.PUBLIC;
    }
  }

  /**
   * Parse RRULE string (e.g., "RRULE;FREQ=DAILY;COUNT=10") into object format
   */
  private parseRRuleString(rruleString: string): any {
    try {
      // Remove 'RRULE' prefix if present
      let rule = rruleString;
      if (rule.startsWith('RRULE:') || rule.startsWith('RRULE;')) {
        rule = rule.substring(6);
      }

      // Parse semicolon-separated key=value pairs
      const parts = rule.split(';');
      const rruleObj: any = {};

      for (const part of parts) {
        const [key, value] = part.split('=');
        if (!key || !value) continue;

        switch (key.toUpperCase()) {
          case 'FREQ':
            rruleObj.freq = value.toUpperCase();
            break;
          case 'INTERVAL':
            rruleObj.interval = parseInt(value, 10);
            break;
          case 'COUNT':
            rruleObj.count = parseInt(value, 10);
            break;
          case 'UNTIL':
            // Parse date string (format: 20240101T000000Z)
            rruleObj.until = new Date(
              value.substring(0, 4) +
                '-' +
                value.substring(4, 6) +
                '-' +
                value.substring(6, 8) +
                'T' +
                value.substring(9, 11) +
                ':' +
                value.substring(11, 13) +
                ':' +
                value.substring(13, 15) +
                'Z'
            );
            break;
          case 'BYDAY':
            rruleObj.byDay = value.split(',');
            break;
          case 'BYMONTHDAY':
            rruleObj.byMonthDay = value.split(',').map((d) => parseInt(d, 10));
            break;
          case 'BYMONTH':
            rruleObj.byMonth = value.split(',').map((m) => parseInt(m, 10));
            break;
          case 'BYSETPOS':
            rruleObj.bySetPos = parseInt(value, 10);
            break;
        }
      }

      return rruleObj;
    } catch (_error) {
      logger.warn(`Failed to parse RRULE string: ${rruleString}`);
      return { freq: ICalEventRepeatingFreq.DAILY }; // Fallback
    }
  }
}
