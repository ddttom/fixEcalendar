import { PSTFile } from 'pst-extractor';
import { PSTParsingError } from '../utils/error-handler';
import { logger } from '../utils/logger';
import { CALENDAR_FOLDER_NAMES, CALENDAR_CONTAINER_CLASS } from '../config/constants';

export class PSTParser {
  private pstFile: any = null;
  private filePath: string = '';

  async open(filePath: string): Promise<void> {
    try {
      logger.debug(`Opening PST file: ${filePath}`);
      this.filePath = filePath;
      this.pstFile = new PSTFile(filePath);
      logger.debug('PST file opened successfully');
    } catch (error) {
      const errorMsg = (error as Error).message;

      // Provide more helpful error messages for common issues
      if (errorMsg.includes('findBtreeItem')) {
        throw new PSTParsingError(
          `Failed to open PST file: ${filePath}\n` +
            `The PST file may be corrupted, encrypted, or in an unsupported format.\n` +
            `Error details: ${errorMsg}`,
          error as Error
        );
      }

      throw new PSTParsingError(
        `Failed to open PST file: ${filePath}\n` + `Error: ${errorMsg}`,
        error as Error
      );
    }
  }

  async getCalendarFolder(): Promise<any> {
    if (!this.pstFile) {
      throw new PSTParsingError('PST file not opened. Call open() first.');
    }

    try {
      logger.debug('Searching for calendar folders...');
      const rootFolder = this.pstFile.getRootFolder();
      const folders = this.findAllCalendarFolders(rootFolder);

      if (folders.length === 0) {
        return null;
      }

      // Return the folder with the most entries (prefer non-empty folders)
      const folderWithMostEntries = folders.reduce((best, current) => {
        const currentCount = current.contentCount || 0;
        const bestCount = best.contentCount || 0;
        return currentCount > bestCount ? current : best;
      });

      logger.info(
        `Selected calendar folder: "${folderWithMostEntries.displayName}" with ${folderWithMostEntries.contentCount} entries`
      );

      if (folders.length > 1) {
        logger.info(`Found ${folders.length} calendar folders total, processing the one with most entries`);
      }

      return folderWithMostEntries;
    } catch (error) {
      throw new PSTParsingError('Failed to locate calendar folder', error as Error);
    }
  }

  async getAllCalendarFolders(): Promise<any[]> {
    if (!this.pstFile) {
      throw new PSTParsingError('PST file not opened. Call open() first.');
    }

    try {
      logger.debug('Searching for all calendar folders...');
      const rootFolder = this.pstFile.getRootFolder();
      const folders = this.findAllCalendarFolders(rootFolder);

      if (folders.length > 0) {
        logger.info(`Found ${folders.length} calendar folder(s) total`);
      }

      return folders;
    } catch (error) {
      throw new PSTParsingError('Failed to locate calendar folders', error as Error);
    }
  }

  private findAllCalendarFolders(folder: any): any[] {
    const calendarFolders: any[] = [];
    const displayName = folder.displayName;
    const containerClass = folder.containerClass;

    // Enhanced logging for troubleshooting
    logger.debug(
      `Checking folder: "${displayName}" (containerClass: "${containerClass || 'undefined'}", contentCount: ${folder.contentCount || 0})`
    );

    // Primary detection: Check container class (Microsoft standard approach)
    let isCalendar = false;
    if (containerClass && typeof containerClass === 'string') {
      const normalizedClass = containerClass.trim().toLowerCase();

      if (normalizedClass.startsWith(CALENDAR_CONTAINER_CLASS.toLowerCase())) {
        logger.info(
          `Found calendar folder via containerClass: "${displayName}" (${containerClass}, ${folder.contentCount || 0} entries)`
        );
        calendarFolders.push(folder);
        isCalendar = true;
      }
    }

    // Fallback detection: Check display name (backward compatibility)
    if (
      !isCalendar &&
      CALENDAR_FOLDER_NAMES.some((name) => displayName?.toLowerCase() === name.toLowerCase())
    ) {
      logger.info(
        `Found calendar folder via display name: "${displayName}"${
          containerClass ? ` (containerClass: ${containerClass})` : ''
        }, ${folder.contentCount || 0} entries`
      );
      calendarFolders.push(folder);
      isCalendar = true;
    }

    // Always search subfolders to find nested calendar folders
    if (folder.hasSubfolders) {
      try {
        const subFolders = folder.getSubFolders();
        for (const subFolder of subFolders) {
          const subResults = this.findAllCalendarFolders(subFolder);
          calendarFolders.push(...subResults);
        }
      } catch (error) {
        logger.warn(`Error searching subfolders of "${displayName}":`, error);
      }
    }

    return calendarFolders;
  }

  getFilePath(): string {
    return this.filePath;
  }

  getPSTFile(): any {
    return this.pstFile;
  }

  close(): void {
    this.pstFile = null;
    logger.debug('PST file closed');
  }
}
