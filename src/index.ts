#!/usr/bin/env node

import { Command } from 'commander';
import * as path from 'path';
import * as fs from 'fs';
import fg from 'fast-glob';
import { PSTParser } from './parser/pst-parser';
import { CalendarExtractor } from './parser/calendar-extractor';
import { ICalConverter } from './converter/ical-converter';
import { validatePSTFile, validateOutputPath, ensureDirectoryExists } from './utils/validators';
import { logger } from './utils/logger';
import { APP_NAME, APP_VERSION } from './config/constants';
import type { ExtractionOptions } from './parser/types';
import type { ConversionOptions, BatchConversionResult } from './converter/types';

interface CLIOptions {
  name?: string;
  timezone?: string;
  output?: string;
  outputDir?: string;
  merge?: boolean;
  startDate?: string;
  endDate?: string;
  noRecurring?: boolean;
  includePrivate?: boolean;
  verbose?: boolean;
}

async function convertSingleFile(
  inputPath: string,
  outputPath: string,
  options: CLIOptions
): Promise<void> {
  logger.info(`Processing: ${inputPath}`);

  // Validate input
  validatePSTFile(inputPath);

  // Parse extraction options
  const extractOptions: ExtractionOptions = {
    includeRecurring: !options.noRecurring,
    includePrivate: options.includePrivate,
  };

  if (options.startDate) {
    extractOptions.dateRangeStart = new Date(options.startDate);
  }

  if (options.endDate) {
    extractOptions.dateRangeEnd = new Date(options.endDate);
  }

  // Open PST file
  const parser = new PSTParser();
  await parser.open(inputPath);

  try {
    // Find calendar folder
    const calendarFolder = await parser.getCalendarFolder();

    if (!calendarFolder) {
      throw new Error('No calendar folder found in PST file');
    }

    // Extract calendar entries
    const extractor = new CalendarExtractor();
    const entries = await extractor.extractFromFolder(calendarFolder, extractOptions);

    if (entries.length === 0) {
      logger.warn('No calendar entries found');
      return;
    }

    logger.info(`Found ${entries.length} calendar entries`);

    // Convert to iCal
    const converter = new ICalConverter();
    const conversionOptions: ConversionOptions = {
      calendarName: options.name,
      timezone: options.timezone,
    };

    const calendar = converter.convert(entries, conversionOptions);

    // Ensure output directory exists
    const outputDir = path.dirname(outputPath);
    ensureDirectoryExists(outputDir);

    // Validate output path
    validateOutputPath(outputPath);

    // Save to file
    await converter.saveToFile(calendar, outputPath);

    logger.success(`✓ Conversion complete: ${outputPath}`);
  } finally {
    parser.close();
  }
}

async function convertMultipleFiles(
  inputPaths: string[],
  options: CLIOptions
): Promise<void> {
  logger.info(`Processing ${inputPaths.length} PST files...`);

  const results: BatchConversionResult = {
    totalFiles: inputPaths.length,
    successfulFiles: 0,
    failedFiles: 0,
    totalEntries: 0,
    results: new Map(),
    errors: new Map(),
  };

  const calendars = [];

  for (let i = 0; i < inputPaths.length; i++) {
    const inputPath = inputPaths[i];
    const fileName = path.basename(inputPath, '.pst');

    logger.info(`[${i + 1}/${inputPaths.length}] Processing: ${fileName}`);

    try {
      // Determine output path
      let outputPath: string;

      if (options.merge) {
        // In merge mode, we'll combine all calendars later
        outputPath = path.join(options.outputDir || '.', `${fileName}.ics`);
      } else if (options.outputDir) {
        outputPath = path.join(options.outputDir, `${fileName}.ics`);
      } else {
        // Place output next to input file
        outputPath = inputPath.replace(/\.pst$/i, '.ics');
      }

      // Validate input
      validatePSTFile(inputPath);

      // Parse extraction options
      const extractOptions: ExtractionOptions = {
        includeRecurring: !options.noRecurring,
        includePrivate: options.includePrivate,
      };

      if (options.startDate) {
        extractOptions.dateRangeStart = new Date(options.startDate);
      }

      if (options.endDate) {
        extractOptions.dateRangeEnd = new Date(options.endDate);
      }

      // Open PST file
      const parser = new PSTParser();
      await parser.open(inputPath);

      try {
        // Find calendar folder
        const calendarFolder = await parser.getCalendarFolder();

        if (!calendarFolder) {
          throw new Error('No calendar folder found in PST file');
        }

        // Extract calendar entries
        const extractor = new CalendarExtractor();
        const entries = await extractor.extractFromFolder(calendarFolder, extractOptions);

        if (entries.length === 0) {
          logger.warn(`  No calendar entries found in ${fileName}`);
          results.successfulFiles++;
          continue;
        }

        logger.info(`  Found ${entries.length} calendar entries`);

        // Convert to iCal
        const converter = new ICalConverter();
        const conversionOptions: ConversionOptions = {
          calendarName: options.name || fileName,
          timezone: options.timezone,
        };

        const calendar = converter.convert(entries, conversionOptions);

        if (options.merge) {
          // Store calendar for later merging
          calendars.push(calendar);
        } else {
          // Save individual file
          ensureDirectoryExists(path.dirname(outputPath));
          await converter.saveToFile(calendar, outputPath);
          logger.success(`  ✓ Saved: ${outputPath}`);
        }

        results.successfulFiles++;
        results.totalEntries += entries.length;
        results.results.set(inputPath, {
          success: true,
          entryCount: entries.length,
          outputPath,
        });
      } finally {
        parser.close();
      }
    } catch (error) {
      results.failedFiles++;
      results.errors.set(inputPath, error as Error);
      logger.error(`  ✗ Failed to process ${fileName}: ${(error as Error).message}`);
    }
  }

  // Handle merge mode
  if (options.merge && calendars.length > 0) {
    logger.info('Merging calendars...');
    const converter = new ICalConverter();
    const mergedCalendar = converter.merge(calendars);

    const outputPath = options.outputDir
      ? path.join(options.outputDir, 'merged-calendar.ics')
      : 'merged-calendar.ics';

    ensureDirectoryExists(path.dirname(outputPath));
    await converter.saveToFile(mergedCalendar, outputPath);
    logger.success(`✓ Merged calendar saved: ${outputPath}`);
  }

  // Print summary
  logger.info('\n=== Conversion Summary ===');
  logger.info(`Total files: ${results.totalFiles}`);
  logger.info(`Successful: ${results.successfulFiles}`);
  logger.info(`Failed: ${results.failedFiles}`);
  logger.info(`Total entries: ${results.totalEntries}`);

  if (results.failedFiles > 0) {
    logger.warn('\nFailed files:');
    results.errors.forEach((error, filePath) => {
      logger.warn(`  - ${path.basename(filePath)}: ${error.message}`);
    });
    process.exit(1);
  }
}

async function main() {
  const program = new Command();

  program
    .name(APP_NAME)
    .description('Convert calendar entries from Microsoft Outlook PST files to iCalendar format')
    .version(APP_VERSION);

  program
    .argument('[inputs...]', 'PST file(s) or glob pattern (e.g., "*.pst", "archive/*.pst")')
    .option('-o, --output <path>', 'Output iCal file path (for single file conversion)')
    .option('-d, --output-dir <dir>', 'Output directory (for multiple files)')
    .option('-m, --merge', 'Merge all PST files into a single calendar')
    .option('-n, --name <name>', 'Calendar name')
    .option('-t, --timezone <tz>', 'Timezone (e.g., "America/New_York", "UTC")')
    .option('--start-date <date>', 'Filter: start date (ISO format: YYYY-MM-DD)')
    .option('--end-date <date>', 'Filter: end date (ISO format: YYYY-MM-DD)')
    .option('--no-recurring', 'Exclude recurring appointments')
    .option('--include-private', 'Include private appointments')
    .option('-v, --verbose', 'Verbose logging')
    .action(async (inputs: string[], options: CLIOptions) => {
      try {
        // Set verbose mode
        if (options.verbose) {
          process.env.DEBUG = '1';
        }

        // Validate inputs
        if (!inputs || inputs.length === 0) {
          program.error('Error: No input files specified');
        }

        // Expand glob patterns
        let inputFiles: string[] = [];
        for (const input of inputs) {
          // Check if it's a glob pattern
          if (input.includes('*') || input.includes('?')) {
            const matches = await fg(input, { onlyFiles: true });
            inputFiles.push(...matches);
          } else {
            inputFiles.push(input);
          }
        }

        // Remove duplicates
        inputFiles = [...new Set(inputFiles)];

        if (inputFiles.length === 0) {
          program.error('Error: No PST files found');
        }

        // Validate output options
        if (inputFiles.length === 1 && !options.merge) {
          // Single file conversion
          const outputPath = options.output || inputFiles[0].replace(/\.pst$/i, '.ics');
          await convertSingleFile(inputFiles[0], outputPath, options);
        } else {
          // Multiple files
          if (options.output) {
            logger.warn('Warning: --output option ignored for multiple files. Use --output-dir instead.');
          }

          // Ensure output directory is specified for non-merge batch processing
          if (!options.merge && !options.outputDir) {
            logger.info('No output directory specified. Output files will be created next to input files.');
          }

          await convertMultipleFiles(inputFiles, options);
        }
      } catch (error) {
        logger.error('Conversion failed:', error as Error);
        process.exit(1);
      }
    });

  program.parse();
}

// Run CLI
if (require.main === module) {
  main();
}

export { convertSingleFile, convertMultipleFiles };
