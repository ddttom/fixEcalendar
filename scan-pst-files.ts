#!/usr/bin/env ts-node

/**
 * PST File Scanner
 * Scans all PST files in a directory and checks if they can be opened
 * Reports which files are valid and which have issues
 */

import * as fs from 'fs';
import * as path from 'path';
import { PSTFile } from 'pst-extractor';

interface PSTFileStatus {
  filePath: string;
  fileName: string;
  size: number;
  status: 'success' | 'error';
  error?: string;
  calendarFolder?: string;
  itemCount?: number;
}

function formatBytes(bytes: number): string {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
}

function scanPSTFile(filePath: string): PSTFileStatus {
  const fileName = path.basename(filePath);
  const stats = fs.statSync(filePath);

  console.log(`\n📄 Scanning: ${fileName}`);
  console.log(`   Size: ${formatBytes(stats.size)}`);

  const result: PSTFileStatus = {
    filePath,
    fileName,
    size: stats.size,
    status: 'success',
  };

  try {
    // Try to open the PST file
    const pstFile = new PSTFile(filePath);
    console.log(`   ✅ File opened successfully`);

    // Try to get root folder
    try {
      const rootFolder = pstFile.getRootFolder();
      console.log(`   ✅ Root folder accessible: "${rootFolder.displayName}"`);

      // Try to find calendar folder
      const calendarFolder = findCalendarFolder(rootFolder);
      if (calendarFolder) {
        result.calendarFolder = calendarFolder.displayName;
        result.itemCount = calendarFolder.contentCount;
        console.log(`   ✅ Calendar folder found: "${calendarFolder.displayName}"`);
        console.log(`   📊 Calendar items: ${calendarFolder.contentCount}`);
      } else {
        console.log(`   ⚠️  No calendar folder found`);
        result.error = 'No calendar folder found';
      }
    } catch (error) {
      const errorMsg = (error as Error).message;
      result.status = 'error';
      result.error = `Root folder error: ${errorMsg}`;
      console.log(`   ❌ Error accessing root folder: ${errorMsg}`);
    }
  } catch (error) {
    const errorMsg = (error as Error).message;
    result.status = 'error';
    result.error = errorMsg;

    if (errorMsg.includes('findBtreeItem')) {
      console.log(`   ❌ File structure error (possibly corrupted)`);
      console.log(`   💡 Try repairing with SCANPST.EXE`);
    } else if (errorMsg.includes('password') || errorMsg.includes('encrypted')) {
      console.log(`   ❌ File is password-protected or encrypted`);
    } else {
      console.log(`   ❌ Error opening file: ${errorMsg}`);
    }
  }

  return result;
}

function findCalendarFolder(folder: any): any {
  const CALENDAR_FOLDER_NAMES = [
    'Calendar',
    'Kalender',
    'Calendrier',
    'Calendario',
  ];

  // Check if current folder is a calendar folder
  const displayName = folder.displayName;
  if (CALENDAR_FOLDER_NAMES.some(name =>
    displayName?.toLowerCase() === name.toLowerCase()
  )) {
    return folder;
  }

  // Recursively search subfolders
  if (folder.hasSubfolders) {
    try {
      const subFolders = folder.getSubFolders();
      for (const subFolder of subFolders) {
        const result = findCalendarFolder(subFolder);
        if (result) {
          return result;
        }
      }
    } catch (error) {
      // Ignore subfolder errors
    }
  }

  return null;
}

function scanDirectory(dirPath: string): PSTFileStatus[] {
  const line = '='.repeat(70);
  console.log(`\n${line}`);
  console.log(`🔍 Scanning PST files in: ${dirPath}`);
  console.log(line);

  if (!fs.existsSync(dirPath)) {
    console.error(`\n❌ Directory not found: ${dirPath}`);
    process.exit(1);
  }

  const files = fs.readdirSync(dirPath);
  const pstFiles = files.filter(f => f.toLowerCase().endsWith('.pst'));

  if (pstFiles.length === 0) {
    console.log(`\n⚠️  No PST files found in directory`);
    return [];
  }

  console.log(`\n📁 Found ${pstFiles.length} PST file(s)\n`);

  const results: PSTFileStatus[] = [];

  for (let i = 0; i < pstFiles.length; i++) {
    const filePath = path.join(dirPath, pstFiles[i]);
    const result = scanPSTFile(filePath);
    results.push(result);
  }

  return results;
}

function printSummary(results: PSTFileStatus[]): void {
  const line = '='.repeat(70);
  const thinLine = '-'.repeat(70);

  console.log(`\n\n${line}`);
  console.log(`📊 SCAN SUMMARY`);
  console.log(`${line}\n`);

  const successful = results.filter(r => r.status === 'success' && !r.error);
  const withCalendar = results.filter(r => r.calendarFolder);
  const errors = results.filter(r => r.status === 'error');
  const noCalendar = results.filter(r => r.status === 'success' && r.error === 'No calendar folder found');

  console.log(`Total PST files scanned: ${results.length}`);
  console.log(`✅ Successfully opened: ${successful.length}`);
  console.log(`📅 With calendar data: ${withCalendar.length}`);
  console.log(`⚠️  Without calendar: ${noCalendar.length}`);
  console.log(`❌ Errors/Corrupted: ${errors.length}`);

  if (withCalendar.length > 0) {
    console.log(`\n${thinLine}`);
    console.log(`✅ FILES WITH CALENDAR DATA (${withCalendar.length}):`);
    console.log(`${thinLine}\n`);

    let totalItems = 0;
    withCalendar.forEach(file => {
      console.log(`✓ ${file.fileName}`);
      console.log(`  Size: ${formatBytes(file.size)}`);
      console.log(`  Calendar: ${file.calendarFolder}`);
      console.log(`  Items: ${file.itemCount}`);
      console.log();
      totalItems += file.itemCount || 0;
    });

    console.log(`📊 Total calendar items across all files: ${totalItems}`);
  }

  if (noCalendar.length > 0) {
    console.log(`\n${thinLine}`);
    console.log(`⚠️  FILES WITHOUT CALENDAR DATA (${noCalendar.length}):`);
    console.log(`${thinLine}\n`);

    noCalendar.forEach(file => {
      console.log(`○ ${file.fileName} (${formatBytes(file.size)})`);
    });
  }

  if (errors.length > 0) {
    console.log(`\n${thinLine}`);
    console.log(`❌ FILES WITH ERRORS (${errors.length}):`);
    console.log(`${thinLine}\n`);

    errors.forEach(file => {
      console.log(`✗ ${file.fileName}`);
      console.log(`  Size: ${formatBytes(file.size)}`);
      console.log(`  Error: ${file.error}`);
      console.log();
    });

    console.log(`💡 TIP: Corrupted files can often be repaired with:`);
    console.log(`   • SCANPST.EXE (Windows - Inbox Repair Tool)`);
    console.log(`   • Re-export from Outlook`);
    console.log(`   • See PST_TROUBLESHOOTING.md for details`);
  }

  console.log(`\n${line}`);
  console.log(`🎯 NEXT STEPS`);
  console.log(`${line}\n`);

  if (withCalendar.length > 0) {
    console.log(`To process the working files with fixECalendar:\n`);

    const workingFiles = withCalendar.map(f => path.basename(f.filePath)).join('" "');

    console.log(`# Process all working files:`);
    console.log(`fixECalendar "${workingFiles}"\n`);

    console.log(`# Check database stats:`);
    console.log(`fixECalendar --db-stats\n`);

    console.log(`# Export to iCal:`);
    console.log(`fixECalendar --output calendar.ics\n`);
  }

  if (errors.length > 0) {
    console.log(`Files with errors should be repaired before processing.`);
    console.log(`See PST_TROUBLESHOOTING.md for repair instructions.\n`);
  }
}

// Main execution
const args = process.argv.slice(2);
const directory = args[0] || './prod';

try {
  const results = scanDirectory(directory);
  printSummary(results);
} catch (error) {
  console.error(`\n❌ Fatal error: ${(error as Error).message}`);
  process.exit(1);
}
