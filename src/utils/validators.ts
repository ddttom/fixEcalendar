import * as fs from 'fs';
import * as path from 'path';
import { ValidationError } from './error-handler';

export function validatePSTFile(filePath: string): void {
  if (!fs.existsSync(filePath)) {
    throw new ValidationError(`PST file not found: ${filePath}`);
  }

  const ext = path.extname(filePath).toLowerCase();
  if (ext !== '.pst') {
    throw new ValidationError(`Invalid file extension: ${ext}. Expected .pst`);
  }

  const stats = fs.statSync(filePath);
  if (!stats.isFile()) {
    throw new ValidationError(`Not a file: ${filePath}`);
  }

  if (stats.size === 0) {
    throw new ValidationError(`PST file is empty: ${filePath}`);
  }
}

export function validateOutputPath(filePath: string): void {
  const dir = path.dirname(filePath);

  if (!fs.existsSync(dir)) {
    throw new ValidationError(`Output directory does not exist: ${dir}`);
  }

  const ext = path.extname(filePath).toLowerCase();
  if (ext !== '.ics') {
    throw new ValidationError(`Invalid output file extension: ${ext}. Expected .ics`);
  }
}

export function ensureDirectoryExists(dirPath: string): void {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}
