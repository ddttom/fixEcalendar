export class PSTParsingError extends Error {
  constructor(message: string, public cause?: Error) {
    super(message);
    this.name = 'PSTParsingError';
    Error.captureStackTrace(this, this.constructor);
  }
}

export class ConversionError extends Error {
  constructor(message: string, public cause?: Error) {
    super(message);
    this.name = 'ConversionError';
    Error.captureStackTrace(this, this.constructor);
  }
}

export class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ValidationError';
    Error.captureStackTrace(this, this.constructor);
  }
}
