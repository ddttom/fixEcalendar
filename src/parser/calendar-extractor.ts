import { CalendarEntry, ExtractionOptions, Attendee } from './types';
import { PSTParsingError } from '../utils/error-handler';
import { logger } from '../utils/logger';
// Import RecurrencePattern from pst-extractor for detailed recurrence parsing
import { RecurrencePattern } from 'pst-extractor/dist/RecurrencePattern.class';

export class CalendarExtractor {
  async extractFromFolder(folder: any, options: ExtractionOptions = {}): Promise<CalendarEntry[]> {
    const entries: CalendarEntry[] = [];
    const errors: string[] = [];

    try {
      logger.info(`Extracting calendar entries from folder: ${folder.displayName}`);

      if (folder.contentCount === 0) {
        logger.warn('Calendar folder is empty');
        return entries;
      }

      logger.debug(`Processing ${folder.contentCount} items...`);

      // Iterate through all items in the folder
      let item = folder.getNextChild();
      let count = 0;

      while (item) {
        try {
          // Check if this is an appointment
          if (this.isAppointment(item)) {
            const appointment = item as any;
            const entry = this.parseAppointment(appointment, options);

            if (entry) {
              entries.push(entry);
              count++;
            }
          }
        } catch (error) {
          const errorMsg = `Error parsing item: ${(error as Error).message}`;
          errors.push(errorMsg);
          logger.warn(errorMsg);
        }

        item = folder.getNextChild();
      }

      logger.info(`Successfully extracted ${count} calendar entries`);
      if (errors.length > 0) {
        logger.warn(`Encountered ${errors.length} errors during extraction`);
      }

      return entries;
    } catch (error) {
      throw new PSTParsingError('Failed to extract calendar entries', error as Error);
    }
  }

  private isAppointment(item: any): boolean {
    // Check if item is an appointment by checking message class
    const messageClass = item.messageClass;
    return (
      messageClass &&
      (messageClass.startsWith('IPM.Appointment') || messageClass.startsWith('IPM.Schedule'))
    );
  }

  private parseAppointment(appointment: any, options: ExtractionOptions): CalendarEntry | null {
    try {
      // Get basic properties
      const subject = appointment.subject || '(No Subject)';
      const startTime = appointment.startTime;
      const endTime = appointment.endTime;

      // Validate required fields
      if (!startTime || !endTime) {
        logger.warn(`Skipping appointment "${subject}": missing start or end time`);
        return null;
      }

      // Apply date range filter if specified
      if (options.dateRangeStart && startTime < options.dateRangeStart) {
        return null;
      }
      if (options.dateRangeEnd && startTime > options.dateRangeEnd) {
        return null;
      }

      // Check for recurring appointments
      const isRecurring = appointment.isRecurring || false;
      if (isRecurring && options.includeRecurring === false) {
        logger.debug(`Skipping recurring appointment: ${subject}`);
        return null;
      }

      // Check sensitivity/privacy
      const sensitivity = this.mapSensitivity(appointment.sensitivity);
      if (sensitivity === 'private' && options.includePrivate === false) {
        logger.debug(`Skipping private appointment: ${subject}`);
        return null;
      }

      // Extract all properties
      const entry: CalendarEntry = {
        subject,
        startTime,
        endTime,
        location: appointment.location || undefined,
        description: this.extractDescription(appointment),
        organizer: appointment.organizer || undefined,
        attendees: this.extractAttendees(appointment),
        isAllDay: this.isAllDayEvent(appointment),
        isRecurring,
        recurrencePattern: isRecurring ? this.extractRecurrencePattern(appointment) : undefined,
        reminder: this.extractReminder(appointment),
        importance: this.mapImportance(appointment.importance),
        busyStatus: this.mapBusyStatus(appointment.busyStatus),
        sensitivity,
        uid: this.generateUID(appointment),
      };

      logger.debug(`Parsed appointment: ${subject}`);
      return entry;
    } catch (error) {
      logger.warn(`Error parsing appointment: ${(error as Error).message}`);
      return null;
    }
  }

  private extractDescription(appointment: any): string | undefined {
    // Try to get body text
    let description = appointment.body;

    // If no plain text body, try HTML body
    if (!description || description.trim() === '') {
      description = appointment.bodyHTML;

      // Strip HTML tags for plain text (basic implementation)
      if (description) {
        description = description
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .trim();
      }
    }

    return description || undefined;
  }

  private extractAttendees(appointment: any): Attendee[] | undefined {
    const attendees: Attendee[] = [];

    try {
      const recipientCount = appointment.numberOfRecipients;

      for (let i = 0; i < recipientCount; i++) {
        try {
          const recipient: any = appointment.getRecipient(i);
          const attendee: Attendee = {
            name: recipient.displayName || undefined,
            email: recipient.emailAddress || undefined,
            type: this.mapRecipientType(recipient.recipientType),
          };

          if (attendee.name || attendee.email) {
            attendees.push(attendee);
          }
        } catch (error) {
          logger.debug(`Error extracting recipient ${i}: ${(error as Error).message}`);
        }
      }
    } catch (error) {
      logger.debug(`Error extracting attendees: ${(error as Error).message}`);
    }

    return attendees.length > 0 ? attendees : undefined;
  }

  private mapRecipientType(type: number): 'required' | 'optional' | 'resource' {
    // MAPI recipient types
    // 1 = TO (required)
    // 2 = CC (optional)
    // 3 = BCC (optional)
    switch (type) {
      case 1:
        return 'required';
      case 2:
      case 3:
        return 'optional';
      default:
        return 'required';
    }
  }

  private isAllDayEvent(appointment: any): boolean {
    // Check if the appointment is marked as an all-day event
    // This can be determined by the subType property or by checking if
    // start and end times are at midnight
    try {
      const subType = (appointment as any).subType;
      if (subType === 1) {
        return true;
      }

      // Alternative: check if times are exactly at midnight
      const start = appointment.startTime;
      const end = appointment.endTime;

      if (start && end) {
        const isStartMidnight =
          start.getHours() === 0 && start.getMinutes() === 0 && start.getSeconds() === 0;
        const isEndMidnight =
          end.getHours() === 0 && end.getMinutes() === 0 && end.getSeconds() === 0;

        if (isStartMidnight && isEndMidnight) {
          return true;
        }
      }
    } catch (error) {
      logger.debug(`Error checking all-day status: ${(error as Error).message}`);
    }

    return false;
  }

  private extractRecurrencePattern(appointment: any): string | undefined {
    try {
      // Try to get the recurrence structure buffer
      const recurrenceStructure = appointment.recurrenceStructure;

      if (recurrenceStructure && Buffer.isBuffer(recurrenceStructure)) {
        // Parse the binary recurrence structure using RecurrencePattern class
        logger.debug(`Parsing recurrence structure for: ${appointment.subject}`);
        const pattern = new RecurrencePattern(recurrenceStructure);

        // Log the parsed pattern for debugging
        logger.debug(`Parsed recurrence pattern: ${JSON.stringify(pattern.toJSON())}`);

        // Build RRULE from parsed pattern
        return this.buildRRuleFromPattern(pattern);
      }

      // Fallback: try using basic recurrence properties
      const recurrenceType = appointment.recurrenceType;
      const recurrencePattern = appointment.recurrencePattern; // Human-readable text

      // If no recurrence type, try to infer from subject line
      let inferredType = recurrenceType;
      if (recurrenceType === undefined) {
        // Infer yearly recurrence from birthday/anniversary keywords
        const subject = (appointment.subject || '').toLowerCase();
        if (subject.includes('birthday') || subject.includes('anniversary')) {
          inferredType = 3; // Yearly
          logger.debug(`Inferred YEARLY recurrence for: ${appointment.subject}`);
        } else {
          return undefined;
        }
      }

      const rruleParts: string[] = ['RRULE'];

      // Map recurrence frequency to RRULE FREQ
      // PST recurrence types: 0=Daily, 1=Weekly, 2=Monthly, 3=Yearly
      switch (inferredType) {
        case 0:
          rruleParts.push('FREQ=DAILY');
          break;
        case 1:
          rruleParts.push('FREQ=WEEKLY');
          break;
        case 2:
          rruleParts.push('FREQ=MONTHLY');
          break;
        case 3:
          rruleParts.push('FREQ=YEARLY');
          break;
        default:
          rruleParts.push('FREQ=DAILY');
      }

      // Parse the recurrence pattern text to extract additional details
      if (recurrencePattern) {
        // Parse day of week from patterns like "every Monday" or "the second Monday"
        const dayMatch = recurrencePattern.match(
          /(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)/i
        );
        if (dayMatch) {
          const dayMap: { [key: string]: string } = {
            monday: 'MO',
            tuesday: 'TU',
            wednesday: 'WE',
            thursday: 'TH',
            friday: 'FR',
            saturday: 'SA',
            sunday: 'SU',
          };
          const day = dayMap[dayMatch[1].toLowerCase()];
          if (day) {
            // Check for ordinal (first, second, third, fourth, last)
            const ordinalMatch = recurrencePattern.match(
              /(first|second|third|fourth|last)\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)/i
            );
            if (ordinalMatch) {
              const ordinalMap: { [key: string]: string } = {
                first: '1',
                second: '2',
                third: '3',
                fourth: '4',
                last: '-1',
              };
              const ordinal = ordinalMap[ordinalMatch[1].toLowerCase()];
              rruleParts.push(`BYDAY=${ordinal}${day}`);
            } else {
              rruleParts.push(`BYDAY=${day}`);
            }
          }
        }

        // Parse day of month from patterns like "day 1 of every 1 month"
        const dayOfMonthMatch = recurrencePattern.match(/day (\d+)/i);
        if (dayOfMonthMatch) {
          rruleParts.push(`BYMONTHDAY=${dayOfMonthMatch[1]}`);
        }

        // Parse interval from patterns like "every 2 weeks"
        const intervalMatch = recurrencePattern.match(/every (\d+) (day|week|month|year)/i);
        if (intervalMatch && parseInt(intervalMatch[1]) > 1) {
          rruleParts.push(`INTERVAL=${intervalMatch[1]}`);
        }
      }

      // Join parts with semicolons
      return rruleParts.join(';');
    } catch (error) {
      logger.debug(`Error extracting recurrence pattern: ${(error as Error).message}`);
      return 'RECURRING'; // Fallback for unparseable recurrences
    }
  }

  private buildRRuleFromPattern(pattern: RecurrencePattern): string {
    const rruleParts: string[] = ['RRULE'];

    // Map frequency - RecurFrequency enum: Daily=8202, Weekly=8203, Monthly=8204, Yearly=8205
    switch (pattern.recurFrequency) {
      case 8202: // RecurFrequency.Daily
        rruleParts.push('FREQ=DAILY');
        break;
      case 8203: // RecurFrequency.Weekly
        rruleParts.push('FREQ=WEEKLY');
        break;
      case 8204: // RecurFrequency.Monthly
        rruleParts.push('FREQ=MONTHLY');
        break;
      case 8205: // RecurFrequency.Yearly
        rruleParts.push('FREQ=YEARLY');
        break;
      default:
        rruleParts.push('FREQ=DAILY');
    }

    // Add interval (period)
    if (pattern.period && pattern.period > 1) {
      rruleParts.push(`INTERVAL=${pattern.period}`);
    }

    // Handle pattern type specific data
    // PatternType: Day=0, Week=1, Month=2, MonthNth=3, MonthEnd=4
    if (pattern.patternTypeSpecific !== null) {
      switch (pattern.patternType) {
        case 1: // Week
          // patternTypeSpecific is WeekSpecific (boolean array for Sun-Sat)
          if (Array.isArray(pattern.patternTypeSpecific)) {
            const days = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
            const weekdays = pattern.patternTypeSpecific as boolean[];
            const selectedDays = days.filter((_, i) => weekdays[i]);
            if (selectedDays.length > 0) {
              rruleParts.push(`BYDAY=${selectedDays.join(',')}`);
            }
          }
          break;

        case 2: // Month (specific day)
        case 4: // MonthEnd
          // patternTypeSpecific is the day number
          if (typeof pattern.patternTypeSpecific === 'number') {
            rruleParts.push(`BYMONTHDAY=${pattern.patternTypeSpecific}`);
          }
          break;

        case 3: // MonthNth (e.g., "second Monday")
          // patternTypeSpecific is MonthNthSpecific
          const monthNth = pattern.patternTypeSpecific as any;
          if (monthNth && monthNth.weekdays && monthNth.nth) {
            const days = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
            const selectedDays = days.filter((_, i) => monthNth.weekdays[i]);
            const nthMap: { [key: number]: string } = {
              1: '1',
              2: '2',
              3: '3',
              4: '4',
              5: '-1', // Last = -1
            };
            const ordinal = nthMap[monthNth.nth] || '1';

            if (selectedDays.length > 0) {
              const byDayParts = selectedDays.map((day) => `${ordinal}${day}`);
              rruleParts.push(`BYDAY=${byDayParts.join(',')}`);
            }
          }
          break;
      }
    }

    // Add end condition
    // EndType: AfterDate=8225, AfterNOccurrences=8226, NeverEnd=8227
    switch (pattern.endType) {
      case 8225: // AfterDate
        if (pattern.endDate) {
          // Format as YYYYMMDD
          const year = pattern.endDate.getFullYear();
          const month = String(pattern.endDate.getMonth() + 1).padStart(2, '0');
          const day = String(pattern.endDate.getDate()).padStart(2, '0');
          rruleParts.push(`UNTIL=${year}${month}${day}`);
        }
        break;

      case 8226: // AfterNOccurrences
        if (pattern.occurrenceCount && pattern.occurrenceCount > 0) {
          rruleParts.push(`COUNT=${pattern.occurrenceCount}`);
        }
        break;

      case 8227: // NeverEnd
        // No additional rule needed
        break;
    }

    return rruleParts.join(';');
  }

  private extractReminder(appointment: any): number | undefined {
    try {
      // Reminder time is typically in minutes before the event
      const reminderSet = (appointment as any).reminderSet;
      const reminderMinutes = (appointment as any).reminderMinutesBeforeStart;

      if (reminderSet && typeof reminderMinutes === 'number') {
        return reminderMinutes;
      }
    } catch (error) {
      logger.debug(`Error extracting reminder: ${(error as Error).message}`);
    }

    return undefined;
  }

  private mapImportance(importance?: number): 'low' | 'normal' | 'high' {
    // MAPI importance values:
    // 0 = Low
    // 1 = Normal
    // 2 = High
    switch (importance) {
      case 0:
        return 'low';
      case 2:
        return 'high';
      default:
        return 'normal';
    }
  }

  private mapBusyStatus(status?: number): 'free' | 'tentative' | 'busy' | 'out-of-office' {
    // Outlook busy status values:
    // 0 = Free
    // 1 = Tentative
    // 2 = Busy
    // 3 = Out of Office
    switch (status) {
      case 0:
        return 'free';
      case 1:
        return 'tentative';
      case 3:
        return 'out-of-office';
      default:
        return 'busy';
    }
  }

  private mapSensitivity(sensitivity?: number): 'normal' | 'personal' | 'private' | 'confidential' {
    // MAPI sensitivity values:
    // 0 = Normal
    // 1 = Personal
    // 2 = Private
    // 3 = Confidential
    switch (sensitivity) {
      case 1:
        return 'personal';
      case 2:
        return 'private';
      case 3:
        return 'confidential';
      default:
        return 'normal';
    }
  }

  private generateUID(appointment: any): string {
    // Generate a unique ID for the event
    // Try to use existing UID if available, otherwise generate one
    try {
      const entryId = (appointment as any).entryId;
      if (entryId) {
        return entryId;
      }
    } catch (_error) {
      // Ignore
    }

    // Fallback: generate UID from subject and start time
    const subject = appointment.subject || 'event';
    const start = appointment.startTime?.getTime() || Date.now();
    return `${subject.replace(/\s+/g, '-')}-${start}`;
  }
}
