# Security Policy

## Supported Versions

We release security updates for the following versions of fixEcalendar:

| Version | Supported          |
| ------- | ------------------ |
| 1.2.x   | :white_check_mark: |
| 1.1.x   | :white_check_mark: |
| 1.0.x   | :white_check_mark: |
| < 1.0   | :x:                |

## Reporting a Vulnerability

If you discover a security vulnerability in fixEcalendar, please report it responsibly:

### Where to Report

**DO NOT** open a public GitHub issue for security vulnerabilities.

Instead, please report security issues by:
1. Opening a private security advisory on GitHub: https://github.com/ddttom/fixEcalendar/security/advisories/new
2. Or emailing the maintainer directly (see GitHub profile for contact)

### What to Include

When reporting a vulnerability, please include:

- **Description**: Clear description of the vulnerability
- **Impact**: What could an attacker accomplish?
- **Steps to Reproduce**: Detailed steps to reproduce the issue
- **Version**: fixEcalendar version affected
- **Environment**: Node.js version, OS, etc.
- **Proof of Concept**: Code or files demonstrating the issue (if safe to share)
- **Suggested Fix**: If you have ideas for remediation

### Response Timeline

- **Initial Response**: Within 48 hours
- **Status Update**: Within 7 days
- **Fix Timeline**: Depends on severity
  - Critical: Within 7 days
  - High: Within 14 days
  - Medium: Within 30 days
  - Low: Next release cycle

### Disclosure Policy

- We follow coordinated disclosure
- We'll work with you to understand and address the issue
- We'll credit you in the security advisory (unless you prefer anonymity)
- Please allow us reasonable time to fix the issue before public disclosure

## Security Considerations

### PST File Handling

**Important Security Notes:**

1. **Sensitive Data**: PST files contain personal calendar data, emails, contacts, and potentially sensitive information.

2. **Local Processing Only**: fixEcalendar processes PST files locally on your machine. No data is sent to external servers.

3. **File Permissions**: Ensure PST files and exported calendars have appropriate file permissions to prevent unauthorized access.

4. **Database Security**: The SQLite database (`.fixecalendar.db`) contains extracted calendar data. Protect this file like you would the original PST files.

5. **Generated Files**: Calendar exports (.ics, .csv) contain personal data. Handle them securely:
   - Don't commit them to version control
   - Share them only through secure channels
   - Delete them when no longer needed

### Best Practices

**For Users:**

- Only process PST files from trusted sources
- Review exported calendar files before sharing
- Use file encryption for sensitive PST files and exports
- Keep fixEcalendar updated to the latest version
- Run fixEcalendar in a secure environment

**For Developers:**

- Never log sensitive calendar data
- Validate all file paths to prevent path traversal
- Sanitize HTML content from PST descriptions
- Use parameterized queries for database operations (already implemented)
- Keep dependencies updated for security patches

### Known Limitations

1. **Password-Protected PST Files**: Not supported. Attempting to process encrypted PST files will fail safely.

2. **Malformed PST Files**: The `pst-extractor` library handles parsing. Malformed files may cause errors but should not pose security risks.

3. **File Size**: Very large PST files (>10GB) may consume significant memory. Monitor resource usage.

## Security Features

### Current Protections

- **Input Validation**: File paths and parameters are validated
- **SQL Injection**: Parameterized queries prevent SQL injection
- **HTML Sanitization**: HTML tags stripped from descriptions
- **Error Handling**: Errors are caught and logged without exposing sensitive data
- **No Network Access**: Tool operates entirely offline
- **Dependency Scanning**: Regular updates for known vulnerabilities

### Future Enhancements

- Additional input validation improvements
- Enhanced error message sanitization
- Optional encryption for database files
- Integrity verification for PST files

## Dependency Security

We regularly monitor our dependencies for security vulnerabilities:

- **pst-extractor**: PST file parsing library
- **ical-generator**: iCalendar generation library
- **better-sqlite3**: SQLite database interface
- **commander**: CLI framework
- **fast-glob**: File pattern matching

Run `npm audit` to check for known vulnerabilities in dependencies.

## Security Updates

When security issues are fixed:

1. We'll publish a security advisory on GitHub
2. We'll release a patch version immediately
3. We'll update this SECURITY.md file if needed
4. We'll notify users through GitHub releases

## Questions?

If you have questions about security but not a specific vulnerability to report, feel free to:
- Open a GitHub discussion
- Contact the maintainer through GitHub

Thank you for helping keep fixEcalendar secure!
