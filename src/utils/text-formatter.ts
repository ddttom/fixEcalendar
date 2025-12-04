/**
 * Text formatting utilities for calendar data
 */

/**
 * Formats a description field for export by:
 * 1. Detecting and removing HTML/CSS junk from malformed Outlook extractions
 * 2. Trimming whitespace from both ends
 * 3. Normalizing line breaks (CRLF to LF)
 * 4. Reducing 3+ consecutive newlines to exactly 2
 * 5. Truncating to 79 characters maximum or returning empty if meaningless
 *
 * @param description - The description text to format
 * @returns Formatted description string, or empty if content is meaningless
 */
export function formatDescription(description: string | undefined): string {
  if (!description) {
    return '';
  }

  // Step 1: Detect and filter out HTML/CSS junk from malformed Outlook extractions
  // Patterns to detect: "false\nfalse\nfalse", "EN-GB", "X-NONE", "/* Style Definitions */"
  const junkPatterns = [
    /^(false\s*)+$/i, // Multiple "false" entries
    /^(EN-GB|X-NONE)\s*$/i, // Language codes
    /\/\*\s*Style\s+Defin?itions?\s*\*\//i, // CSS comments
    /table\.MsoNormal/i, // Microsoft Word HTML artifacts
    /^\s*(true|false)\s*$/i, // Lone boolean strings
  ];

  // Check if description is primarily junk
  const trimmed = description.trim();
  for (const pattern of junkPatterns) {
    if (pattern.test(trimmed)) {
      return ''; // Return empty for pure junk
    }
  }

  // If description starts with junk patterns, return empty
  if (
    trimmed.startsWith('false\n') ||
    trimmed.startsWith('EN-GB') ||
    trimmed.startsWith('X-NONE') ||
    trimmed.includes('/* Style Defin')
  ) {
    return '';
  }

  // Step 2: Trim whitespace from both ends
  let formatted = trimmed;

  // Step 3: Normalize line breaks (CRLF to LF)
  formatted = formatted.replace(/\r\n/g, '\n');

  // Step 4: Replace 3+ consecutive newlines with exactly 2 newlines
  formatted = formatted.replace(/\n{3,}/g, '\n\n');

  // Step 5: Truncate to 79 characters maximum
  if (formatted.length > 79) {
    formatted = formatted.substring(0, 79);
  }

  return formatted;
}
