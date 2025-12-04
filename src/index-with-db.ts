#!/usr/bin/env node

import { Command } from 'commander';
import * as path from 'path';
import fg from 'fast-glob';
import { PSTParser } from './parser/pst-parser';
import { CalendarExtractor } from './parser/calendar-extractor';
import { ICalConverter } from './converter/ical-converter';
import { CalendarDatabase } from './database/calendar-db';
import { validatePSTFile, validateOutputPath, ensureDirectoryExists } from './utils/validators';
import { logger } from './utils/logger';
import { APP_NAME, APP_VERSION } from './config/constants';
import type { ExtractionOptions } from './parser/types';
import type { ConversionOptions } from './converter/types';

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
  database?: string;
  noDedupe?: boolean;
  clearDb?: boolean;
  dbStats?: boolean;
}

async function processWithDatabase(
  inputPath: string,
  database: CalendarDatabase,
  options: CLIOptions
): Promise<{ found: number; added: number; skipped: number }> {
  logger.info(`Processing: ${path.basename(inputPath)}`);

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
    // Find all calendar folders
    const calendarFolders = await parser.getAllCalendarFolders();

    if (calendarFolders.length === 0) {
      throw new Error('No calendar folder found in PST file');
    }

    // Extract calendar entries from all folders
    const extractor = new CalendarExtractor();
    let totalFound = 0;
    let totalAdded = 0;
    let totalSkipped = 0;

    for (let i = 0; i < calendarFolders.length; i++) {
      const folder = calendarFolders[i];
      const folderPrefix =
        calendarFolders.length > 1 ? `[Folder ${i + 1}/${calendarFolders.length}] ` : '';

      logger.info(
        `${folderPrefix}Extracting from: "${folder.displayName}" (${folder.contentCount || 0} entries)`
      );

      const entries = await extractor.extractFromFolder(folder, extractOptions);

      if (entries.length === 0) {
        logger.warn(`${folderPrefix}No calendar entries found in "${folder.displayName}"`);
        continue;
      }

      logger.info(`${folderPrefix}Found ${entries.length} calendar entries`);

      // Add entries to database with deduplication
      let added = 0;
      let skipped = 0;

      logger.info(`${folderPrefix}Storing entries in database...`);
      for (const entry of entries) {
        if (database.addEntry(entry, inputPath)) {
          added++;
          if (added % 100 === 0) {
            logger.debug(`${folderPrefix}Stored ${added}/${entries.length} entries...`);
          }
        } else {
          skipped++;
        }
      }

      totalFound += entries.length;
      totalAdded += added;
      totalSkipped += skipped;

      logger.success(
        `${folderPrefix}✓ Processed "${folder.displayName}": ${added} added, ${skipped} duplicates skipped`
      );
    }

    if (calendarFolders.length > 1) {
      logger.success(
        `✓ Total for ${path.basename(inputPath)}: ${totalAdded} added, ${totalSkipped} duplicates skipped from ${calendarFolders.length} folders`
      );
    }

    // Log to database
    database.logProcessing(inputPath, totalFound, totalAdded, totalSkipped);

    return { found: totalFound, added: totalAdded, skipped: totalSkipped };
  } finally {
    parser.close();
  }
}

async function exportFromDatabase(
  database: CalendarDatabase,
  outputPath: string,
  options: CLIOptions
): Promise<void> {
  logger.info('Exporting calendar from database...');

  // Get entries from database
  let entries;
  if (options.startDate || options.endDate) {
    const startDate = options.startDate ? new Date(options.startDate) : undefined;
    const endDate = options.endDate ? new Date(options.endDate) : undefined;
    entries = database.getEntriesInRange(startDate, endDate);
  } else {
    entries = database.getAllEntries();
  }

  if (entries.length === 0) {
    logger.warn('No entries to export');
    return;
  }

  logger.info(`Exporting ${entries.length} calendar entries...`);

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

  logger.success(`✓ Exported to: ${outputPath}`);
}

async function showDatabaseStats(database: CalendarDatabase): Promise<void> {
  const stats = database.getStats();

  console.log('\n=== Database Statistics ===');
  console.log(`Database: ${database.getPath()}`);
  console.log(`Total entries: ${stats.totalEntries}`);
  console.log(`Source files processed: ${stats.sourceFiles}`);

  if (stats.dateRange.earliest && stats.dateRange.latest) {
    console.log(
      `Date range: ${stats.dateRange.earliest.toISOString()} to ${stats.dateRange.latest.toISOString()}`
    );
  }

  if (stats.entriesBySource.size > 0) {
    console.log('\nEntries by source file:');
    stats.entriesBySource.forEach((count, file) => {
      console.log(`  - ${path.basename(file)}: ${count} entries`);
    });
  }
}

async function main() {
  const program = new Command();

  program
    .name(APP_NAME)
    .description(
      'Convert calendar entries from Microsoft Outlook PST files to iCalendar format\n' +
        'Uses an intermediate database for deduplication and handling large files'
    )
    .version(APP_VERSION);

  program
    .argument('[inputs...]', 'PST file(s) or glob pattern (e.g., "*.pst", "archive/*.pst")')
    .option('-o, --output <path>', 'Output iCal file path')
    .option('-d, --output-dir <dir>', 'Output directory (for multiple files)')
    .option('-m, --merge', 'Merge all PST files into a single calendar')
    .option('-n, --name <name>', 'Calendar name')
    .option('-t, --timezone <tz>', 'Timezone (e.g., "America/New_York", "UTC")')
    .option('--start-date <date>', 'Filter: start date (ISO format: YYYY-MM-DD)')
    .option('--end-date <date>', 'Filter: end date (ISO format: YYYY-MM-DD)')
    .option('--no-recurring', 'Exclude recurring appointments')
    .option('--include-private', 'Include private appointments')
    .option('--database <path>', 'Path to database file (default: .fixecalendar.db)')
    .option('--no-dedupe', 'Disable deduplication (not recommended for large files)')
    .option('--clear-db', 'Clear database before processing')
    .option('--db-stats', 'Show database statistics and exit')
    .option('-v, --verbose', 'Verbose logging')
    .action(async (inputs: string[], options: CLIOptions) => {
      try {
        // Set verbose mode
        if (options.verbose) {
          process.env.DEBUG = '1';
        }

        // Initialize database
        const database = new CalendarDatabase(options.database);

        try {
          // Show stats and exit if requested
          if (options.dbStats) {
            await showDatabaseStats(database);
            database.close();
            return;
          }

          // Clear database if requested
          if (options.clearDb) {
            logger.info('Clearing database...');
            database.clear();
            logger.success('Database cleared');
          }

          // Validate inputs
          if (!inputs || inputs.length === 0) {
            // If no inputs, try to export from existing database
            if (options.output) {
              const stats = database.getStats();
              if (stats.totalEntries === 0) {
                program.error('Error: No input files specified and database is empty');
              }

              await exportFromDatabase(database, options.output, options);
              database.close();
              return;
            }

            program.error('Error: No input files specified');
          }

          // Expand glob patterns
          let inputFiles: string[] = [];
          for (const input of inputs) {
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

          logger.info(`Processing ${inputFiles.length} PST file(s)...`);

          // Process all files into database
          let totalFound = 0;
          let totalAdded = 0;
          let totalSkipped = 0;

          // Track problematic files for end report
          const filesWithErrors: Array<{ file: string; error: string }> = [];
          const filesWithZeroEntries: string[] = [];
          const filesWithOnlyDuplicates: Array<{ file: string; count: number }> = [];

          for (let i = 0; i < inputFiles.length; i++) {
            const inputPath = inputFiles[i];
            const fileName = path.basename(inputPath);
            logger.info(`[${i + 1}/${inputFiles.length}] ${fileName}`);

            try {
              const result = await processWithDatabase(inputPath, database, options);
              totalFound += result.found;
              totalAdded += result.added;
              totalSkipped += result.skipped;

              // Track files with zero entries
              if (result.found === 0) {
                filesWithZeroEntries.push(fileName);
              }
              // Track files where all entries were duplicates
              else if (result.added === 0 && result.found > 0) {
                filesWithOnlyDuplicates.push({ file: fileName, count: result.found });
              }
            } catch (error) {
              const errorMsg = (error as Error).message;
              logger.error(`Failed to process ${inputPath}:`, error as Error);
              filesWithErrors.push({ file: fileName, error: errorMsg });
            }
          }

          // Show summary
          console.log('\n=== Processing Summary ===');
          console.log(`Files processed: ${inputFiles.length}`);
          console.log(`Total entries found: ${totalFound}`);
          console.log(`New entries added: ${totalAdded}`);
          console.log(`Duplicates skipped: ${totalSkipped}`);

          // Show file status report if there are problematic files
          if (
            filesWithErrors.length > 0 ||
            filesWithZeroEntries.length > 0 ||
            filesWithOnlyDuplicates.length > 0
          ) {
            console.log('\n=== File Status Report ===');

            if (filesWithErrors.length > 0) {
              console.log(`\nFiles with errors (${filesWithErrors.length}):`);
              filesWithErrors.forEach(({ file, error }) => {
                console.log(`  - ${file}: ${error}`);
              });
            }

            if (filesWithZeroEntries.length > 0) {
              console.log(`\nFiles with zero entries (${filesWithZeroEntries.length}):`);
              filesWithZeroEntries.forEach((file) => {
                console.log(`  - ${file}`);
              });
            }

            if (filesWithOnlyDuplicates.length > 0) {
              console.log(`\nFiles with only duplicates (${filesWithOnlyDuplicates.length}):`);
              filesWithOnlyDuplicates.forEach(({ file, count }) => {
                console.log(`  - ${file}: ${count} entries were duplicates`);
              });
            }
          }

          // Show database stats
          await showDatabaseStats(database);

          // Export to iCal if output specified
          if (options.output || options.merge) {
            const outputPath =
              options.output || path.join(options.outputDir || '.', 'merged-calendar.ics');

            await exportFromDatabase(database, outputPath, options);
          } else {
            console.log('\nTo export to iCal format, run:');
            console.log(`  ${APP_NAME} --output calendar.ics`);
          }

          // Optimize database
          logger.info('Optimizing database...');
          database.optimize();

          database.close();
        } catch (error) {
          database.close();
          throw error;
        }
      } catch (error) {
        logger.error('Operation failed:', error as Error);
        process.exit(1);
      }
    });

  program.parse();
}

// Run CLI
if (require.main === module) {
  main();
}

export { processWithDatabase, exportFromDatabase, showDatabaseStats };
