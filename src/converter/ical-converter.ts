import ical, { ICalCalendar } from 'ical-generator';
import * as fs from 'fs';
import { CalendarEntry } from '../parser/types';
import { ConversionOptions, ConversionResult } from './types';
import { PropertyMapper } from './property-mapper';
import { ConversionError } from '../utils/error-handler';
import { logger } from '../utils/logger';
import {
  DEFAULT_CALENDAR_NAME,
  DEFAULT_TIMEZONE,
  APP_NAME,
  APP_VERSION,
} from '../config/constants';

export class ICalConverter {
  private mapper: PropertyMapper;

  constructor() {
    this.mapper = new PropertyMapper();
  }

  convert(entries: CalendarEntry[], options: ConversionOptions = {}): ICalCalendar {
    try {
      logger.debug(`Converting ${entries.length} calendar entries to iCal format`);

      const calendar = ical({
        name: options.calendarName || DEFAULT_CALENDAR_NAME,
        prodId: options.productId || `//${APP_NAME}//v${APP_VERSION}//EN`,
        timezone: options.timezone || DEFAULT_TIMEZONE,
      });

      let successCount = 0;
      const errors: string[] = [];

      for (const entry of entries) {
        try {
          const eventData = this.mapper.mapToICalEvent(entry);
          calendar.createEvent(eventData);
          successCount++;
        } catch (error) {
          const errorMsg = `Error converting entry "${entry.subject}": ${(error as Error).message}`;
          errors.push(errorMsg);
          logger.warn(errorMsg);
        }
      }

      logger.info(`Successfully converted ${successCount}/${entries.length} calendar entries`);

      if (errors.length > 0) {
        logger.warn(`Encountered ${errors.length} errors during conversion`);
      }

      return calendar;
    } catch (error) {
      throw new ConversionError('Failed to convert calendar entries to iCal', error as Error);
    }
  }

  async saveToFile(calendar: ICalCalendar, outputPath: string): Promise<ConversionResult> {
    try {
      logger.debug(`Saving iCal to file: ${outputPath}`);

      const calendarString = calendar.toString();

      await fs.promises.writeFile(outputPath, calendarString, 'utf8');

      const entryCount = calendar.events().length;

      logger.success(`Successfully saved ${entryCount} events to ${outputPath}`);

      return {
        success: true,
        entryCount,
        outputPath,
      };
    } catch (error) {
      throw new ConversionError(`Failed to save iCal file: ${outputPath}`, error as Error);
    }
  }

  toString(calendar: ICalCalendar): string {
    return calendar.toString();
  }

  merge(calendars: ICalCalendar[]): ICalCalendar {
    logger.debug(`Merging ${calendars.length} calendars`);

    // Create a new calendar for the merged result
    const mergedCalendar = ical({
      name: 'Merged Calendar',
      prodId: `//${APP_NAME}//v${APP_VERSION}//EN`,
      timezone: DEFAULT_TIMEZONE,
    });

    // Add all events from all calendars
    let totalEvents = 0;
    for (const calendar of calendars) {
      const events = calendar.events();
      for (const event of events) {
        // Get event data and create new event in merged calendar
        const eventData = {
          start: event.start(),
          end: event.end(),
          summary: event.summary(),
          description: event.description(),
          location: event.location(),
          id: event.id(),
          organizer: event.organizer(),
          attendees: event.attendees(),
          allDay: event.allDay(),
          status: event.status(),
          busystatus: event.busystatus(),
          priority: event.priority(),
          class: event.class(),
          alarms: event.alarms(),
        };

        mergedCalendar.createEvent(eventData);
        totalEvents++;
      }
    }

    logger.info(`Merged ${totalEvents} events from ${calendars.length} calendars`);

    return mergedCalendar;
  }
}
