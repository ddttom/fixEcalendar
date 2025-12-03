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
