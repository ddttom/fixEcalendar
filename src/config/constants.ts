export const APP_NAME = 'fixECalendar';
export const APP_VERSION = '1.0.0';
export const DEFAULT_TIMEZONE = 'UTC';
export const DEFAULT_CALENDAR_NAME = 'Exported from PST';

export const CALENDAR_FOLDER_NAMES = [
  'Calendar',
  'Kalender', // German
  'Calendrier', // French
  'Calendario', // Spanish/Italian
];

// Container class value for calendar folders (Microsoft standard)
export const CALENDAR_CONTAINER_CLASS = 'IPF.Appointment';

export const IMPORTANCE_MAP: Record<number, number> = {
  0: 9, // Low → iCal Priority 9
  1: 5, // Normal → iCal Priority 5
  2: 1, // High → iCal Priority 1
};

export const STATUS_MAP: Record<string, 'FREE' | 'BUSY' | 'TENTATIVE' | 'OOF'> = {
  olFree: 'FREE',
  olBusy: 'BUSY',
  olTentative: 'TENTATIVE',
  olOutOfOffice: 'OOF',
};

/**
 * Recurrence pattern validation configuration
 * Used to detect and fix corrupted Microsoft Outlook recurrence patterns
 * that span unreasonably long time periods (e.g., daily for 75+ years)
 */
export const RECURRENCE_VALIDATION_CONFIG = {
  // Maximum reasonable years for each frequency type
  MAX_YEARS_DAILY: 5, // Daily events capped at 5 years
  MAX_YEARS_WEEKLY: 10, // Weekly events capped at 10 years
  MAX_YEARS_MONTHLY: 20, // Monthly events capped at 20 years
  MAX_YEARS_YEARLY: 100, // Yearly events (birthdays) capped at 100 years

  // Detection thresholds for suspicious patterns
  DAILY_INTERVAL_SUGGEST_WEEKLY: 7, // INTERVAL >= 7 suggests weekly pattern
  DAILY_INTERVAL_SUGGEST_MONTHLY: 28, // INTERVAL >= 28 suggests monthly pattern

  // Occurrence count thresholds
  MAX_OCCURRENCE_COUNT_FOR_PREFERENCE: 50, // Prefer COUNT over corrupted UNTIL if occurrenceCount <= 50

  // Behavior flags
  STRIP_SINGLE_OCCURRENCE: true, // Remove recurrence if occurrenceCount === 1
  CAP_SUSPICIOUS_DATES: true, // Cap UNTIL dates that exceed max years
  PREFER_COUNT_FOR_SMALL_OCCURRENCES: true, // Use COUNT instead of corrupted UNTIL for small occurrence counts
  LOG_VALIDATION_WARNINGS: true, // Log all validation changes
};
