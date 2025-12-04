/**
 * Text formatting utilities for calendar data
 */

/**
 * Formats a description field for export by:
 * 1. Trimming whitespace from both ends
 * 2. Normalizing line breaks (CRLF to LF)
 * 3. Reducing 3+ consecutive newlines to exactly 2
 * 4. Truncating to 79 characters maximum
 *
 * @param description - The description text to format
 * @returns Formatted description string
 */
export function formatDescription(description: string | undefined): string {
  if (!description) {
    return '';
  }

  // Step 1: Trim whitespace from both ends
  let formatted = description.trim();

  // Step 2: Normalize line breaks (CRLF to LF)
  formatted = formatted.replace(/\r\n/g, '\n');

  // Step 3: Replace 3+ consecutive newlines with exactly 2 newlines
  formatted = formatted.replace(/\n{3,}/g, '\n\n');

  // Step 4: Truncate to 79 characters maximum
  if (formatted.length > 79) {
    formatted = formatted.substring(0, 79);
  }

  return formatted;
}
