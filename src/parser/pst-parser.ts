import { PSTFile } from 'pst-extractor';
import { PSTParsingError } from '../utils/error-handler';
import { logger } from '../utils/logger';
import { CALENDAR_FOLDER_NAMES } from '../config/constants';

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
        `Failed to open PST file: ${filePath}\n` +
        `Error: ${errorMsg}`,
        error as Error
      );
    }
  }

  async getCalendarFolder(): Promise<any> {
    if (!this.pstFile) {
      throw new PSTParsingError('PST file not opened. Call open() first.');
    }

    try {
      logger.debug('Searching for calendar folder...');
      const rootFolder = this.pstFile.getRootFolder();
      return this.findCalendarFolder(rootFolder);
    } catch (error) {
      throw new PSTParsingError(
        'Failed to locate calendar folder',
        error as Error
      );
    }
  }

  private findCalendarFolder(folder: any): any {
    // Check if current folder is a calendar folder
    const displayName = folder.displayName;
    logger.debug(`Checking folder: ${displayName}`);

    if (CALENDAR_FOLDER_NAMES.some(name =>
      displayName?.toLowerCase() === name.toLowerCase()
    )) {
      logger.info(`Found calendar folder: ${displayName}`);
      return folder;
    }

    // Recursively search subfolders
    if (folder.hasSubfolders) {
      try {
        const subFolders = folder.getSubFolders();
        for (const subFolder of subFolders) {
          const result = this.findCalendarFolder(subFolder);
          if (result) {
            return result;
          }
        }
      } catch (error) {
        logger.warn(`Error searching subfolders of ${displayName}:`, error);
      }
    }

    return null;
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
