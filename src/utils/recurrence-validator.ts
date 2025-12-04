import { logger } from './logger';

/**
 * Configuration for recurrence pattern validation
 */
export interface RecurrenceValidationConfig {
  maxYearsDaily: number; // Max years for daily recurrence (default: 5)
  maxYearsWeekly: number; // Max years for weekly recurrence (default: 10)
  maxYearsMonthly: number; // Max years for monthly recurrence (default: 20)
  maxYearsYearly: number; // Max years for yearly recurrence (default: 100)
  stripSingleOccurrence: boolean; // Strip recurrence if occurrenceCount === 1
  logAllChanges: boolean; // Log all validation changes
}

/**
 * Result of recurrence pattern validation and fixing
 */
export interface ValidationResult {
  modified: boolean;
  action: 'KEPT' | 'STRIPPED' | 'CAPPED';
  newPattern: string | null;
  newYearsSpan?: number;
  reason: string;
}

/**
 * Default configuration for recurrence validation
 */
const DEFAULT_CONFIG: RecurrenceValidationConfig = {
  maxYearsDaily: 5,
  maxYearsWeekly: 10,
  maxYearsMonthly: 20,
  maxYearsYearly: 100,
  stripSingleOccurrence: true,
  logAllChanges: true,
};

/**
 * Microsoft Outlook RecurrencePattern frequency constants
 * From MS-OXOCAL specification
 */
export enum RecurrenceFrequency {
  DAILY = 8202,
  WEEKLY = 8203,
  MONTHLY = 8204,
  YEARLY = 8205,
}

/**
 * Microsoft Outlook RecurrencePattern end type constants
 */
export enum RecurrenceEndType {
  NEVER_END = 8227,
  AFTER_COUNT = 8226,
  AFTER_DATE = 8225,
}

/**
 * Validates and fixes suspicious recurrence patterns
 */
export class RecurrenceValidator {
  private static config: RecurrenceValidationConfig = DEFAULT_CONFIG;

  /**
   * Set custom validation configuration
   */
  static setConfig(config: Partial<RecurrenceValidationConfig>): void {
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * Get maximum reasonable years for a given frequency
   */
  static getMaxYearsForFrequency(frequency: number): number {
    switch (frequency) {
      case RecurrenceFrequency.DAILY:
        return this.config.maxYearsDaily;
      case RecurrenceFrequency.WEEKLY:
        return this.config.maxYearsWeekly;
      case RecurrenceFrequency.MONTHLY:
        return this.config.maxYearsMonthly;
      case RecurrenceFrequency.YEARLY:
        return this.config.maxYearsYearly;
      default:
        return this.config.maxYearsWeekly; // Default to weekly
    }
  }

  /**
   * Calculate years span between two dates
   */
  private static calculateYearsSpan(startTime: Date, endTime: Date): number {
    return endTime.getFullYear() - startTime.getFullYear();
  }

  /**
   * Rebuild RRULE with capped UNTIL date
   */
  private static rebuildRRuleWithCappedUntil(rrule: string, cappedEnd: Date): string {
    const year = cappedEnd.getFullYear();
    const month = String(cappedEnd.getMonth() + 1).padStart(2, '0');
    const day = String(cappedEnd.getDate()).padStart(2, '0');

    return rrule.replace(/UNTIL=\d{8}T\d{6}Z/, `UNTIL=${year}${month}${day}T235959Z`);
  }

  /**
   * Validate recurrence pattern during PST extraction
   * This is called with the RecurrencePattern object from pst-extractor
   */
  static validateRecurrencePattern(
    pattern: any, // RecurrencePattern from pst-extractor
    startTime: Date,
    subject: string,
    rrule: string | null
  ): string | null {
    // If no RRULE was generated, nothing to validate
    if (!rrule) {
      return null;
    }

    // Check if this is a single occurrence event
    if (this.config.stripSingleOccurrence && pattern.occurrenceCount === 1) {
      if (this.config.logAllChanges) {
        logger.warn(
          `Stripping recurrence for "${subject}" (occurrenceCount=1 indicates single event)`
        );
      }
      return null;
    }

    // Check if recurrence has an UNTIL date that needs validation
    if (pattern.endType === RecurrenceEndType.AFTER_DATE && pattern.endDate) {
      const yearsSpan = this.calculateYearsSpan(startTime, pattern.endDate);
      const maxYears = this.getMaxYearsForFrequency(pattern.recurFrequency);

      if (yearsSpan > maxYears) {
        // Cap the UNTIL date to a reasonable span
        const cappedEnd = new Date(startTime);
        cappedEnd.setFullYear(startTime.getFullYear() + maxYears);

        if (this.config.logAllChanges) {
          logger.warn(
            `Capping recurrence for "${subject}": ${yearsSpan} years → ${maxYears} years (UNTIL date too far in future)`
          );
        }

        return this.rebuildRRuleWithCappedUntil(rrule, cappedEnd);
      }
    }

    // Check for suspicious interval patterns (e.g., INTERVAL=7 for DAILY)
    if (pattern.recurFrequency === RecurrenceFrequency.DAILY && pattern.period >= 7) {
      if (this.config.logAllChanges) {
        logger.info(
          `Daily recurrence with INTERVAL=${pattern.period} detected for "${subject}" (may be misclassified weekly pattern)`
        );
      }
    }

    // Pattern is reasonable, return as-is
    return rrule;
  }

  /**
   * Validate and fix existing RRULE string from database/CSV
   * Used by cleanup scripts and CSV import
   */
  static validateAndCap(
    rruleString: string,
    startTime: Date,
    subject?: string
  ): ValidationResult {
    // Extract frequency from RRULE
    const freqMatch = rruleString.match(/FREQ=(DAILY|WEEKLY|MONTHLY|YEARLY)/);
    if (!freqMatch) {
      return {
        modified: false,
        action: 'KEPT',
        newPattern: rruleString,
        reason: 'No frequency found in RRULE',
      };
    }

    const frequency = freqMatch[1];

    // Extract UNTIL date if present
    const untilMatch = rruleString.match(/UNTIL=(\d{8})/);
    if (!untilMatch) {
      // No UNTIL date, pattern is either COUNT-based or infinite (both legitimate)
      return {
        modified: false,
        action: 'KEPT',
        newPattern: rruleString,
        reason: 'No UNTIL date (COUNT-based or infinite recurrence)',
      };
    }

    const untilYear = parseInt(untilMatch[1].substring(0, 4));
    const yearsSpan = untilYear - startTime.getFullYear();

    // Determine max years based on frequency
    let maxYears: number;
    switch (frequency) {
      case 'DAILY':
        maxYears = this.config.maxYearsDaily;
        break;
      case 'WEEKLY':
        maxYears = this.config.maxYearsWeekly;
        break;
      case 'MONTHLY':
        maxYears = this.config.maxYearsMonthly;
        break;
      case 'YEARLY':
        maxYears = this.config.maxYearsYearly;
        break;
      default:
        maxYears = this.config.maxYearsWeekly;
    }

    // Check if years span exceeds maximum
    if (yearsSpan > maxYears) {
      const cappedEnd = new Date(startTime);
      cappedEnd.setFullYear(startTime.getFullYear() + maxYears);
      const cappedYear = cappedEnd.getFullYear();
      const cappedMonth = String(cappedEnd.getMonth() + 1).padStart(2, '0');
      const cappedDay = String(cappedEnd.getDate()).padStart(2, '0');

      const newPattern = rruleString.replace(
        /UNTIL=\d{8}(T\d{6}Z)?/,
        `UNTIL=${cappedYear}${cappedMonth}${cappedDay}T235959Z`
      );

      if (this.config.logAllChanges && subject) {
        logger.warn(
          `Capped ${frequency} recurrence for "${subject}": ${yearsSpan} years → ${maxYears} years`
        );
      }

      return {
        modified: true,
        action: 'CAPPED',
        newPattern,
        newYearsSpan: maxYears,
        reason: `${frequency} recurrence span (${yearsSpan} years) exceeds maximum (${maxYears} years)`,
      };
    }

    // Pattern is reasonable
    return {
      modified: false,
      action: 'KEPT',
      newPattern: rruleString,
      reason: `${frequency} recurrence span (${yearsSpan} years) is reasonable`,
    };
  }

  /**
   * Get frequency enum from RRULE string
   */
  static getFrequencyFromRRule(rrule: string): RecurrenceFrequency | null {
    const match = rrule.match(/FREQ=(DAILY|WEEKLY|MONTHLY|YEARLY)/);
    if (!match) return null;

    switch (match[1]) {
      case 'DAILY':
        return RecurrenceFrequency.DAILY;
      case 'WEEKLY':
        return RecurrenceFrequency.WEEKLY;
      case 'MONTHLY':
        return RecurrenceFrequency.MONTHLY;
      case 'YEARLY':
        return RecurrenceFrequency.YEARLY;
      default:
        return null;
    }
  }

  /**
   * Check if RRULE has suspicious UNTIL=2100 pattern
   */
  static hasSuspiciousUntil(rrule: string, startTime: Date): boolean {
    const untilMatch = rrule.match(/UNTIL=(\d{8})/);
    if (!untilMatch) return false;

    const untilYear = parseInt(untilMatch[1].substring(0, 4));
    const yearsSpan = untilYear - startTime.getFullYear();

    // Check if UNTIL is 2100 and span is > 70 years (likely corrupted)
    return untilYear === 2100 && yearsSpan > 70;
  }
}
