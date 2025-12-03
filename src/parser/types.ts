export interface CalendarEntry {
  subject: string;
  startTime: Date;
  endTime: Date;
  location?: string;
  description?: string;
  organizer?: string;
  attendees?: Attendee[];
  isAllDay: boolean;
  isRecurring: boolean;
  recurrencePattern?: string;
  reminder?: number; // minutes before event
  importance?: 'low' | 'normal' | 'high';
  busyStatus?: 'free' | 'tentative' | 'busy' | 'out-of-office';
  sensitivity?: 'normal' | 'personal' | 'private' | 'confidential';
  uid?: string;
}

export interface Attendee {
  name?: string;
  email?: string;
  type?: 'required' | 'optional' | 'resource';
}

export interface ExtractionOptions {
  includeRecurring?: boolean;
  dateRangeStart?: Date;
  dateRangeEnd?: Date;
  includePrivate?: boolean;
}

export interface PSTFileMetadata {
  fileName: string;
  fileSize: number;
  messageCount?: number;
}
