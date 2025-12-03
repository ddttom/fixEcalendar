import { CalendarEntry } from '../parser/types';

export interface ConversionOptions {
  calendarName?: string;
  timezone?: string;
  productId?: string;
}

export interface ConversionResult {
  success: boolean;
  entryCount: number;
  outputPath: string;
  errors?: string[];
}

export interface BatchConversionResult {
  totalFiles: number;
  successfulFiles: number;
  failedFiles: number;
  totalEntries: number;
  results: Map<string, ConversionResult>;
  errors: Map<string, Error>;
}

export { CalendarEntry };
