# Contributing to fixEcalendar

Thank you for your interest in contributing to fixEcalendar! This document provides guidelines and instructions for contributing to the project.

## Code of Conduct

By participating in this project, you agree to maintain a respectful and inclusive environment for all contributors.

## Getting Started

### Prerequisites

- Node.js 16 or higher
- npm (comes with Node.js)
- Git

### Development Setup

1. Fork the repository on GitHub
2. Clone your fork locally:
   ```bash
   git clone https://github.com/YOUR_USERNAME/fixEcalendar.git
   cd fixEcalendar
   ```

3. Install dependencies:
   ```bash
   npm install
   ```

4. Create a branch for your changes:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## Development Workflow

### Building the Project

```bash
npm run build
```

This compiles TypeScript to JavaScript in the `dist/` directory.

### Running in Development Mode

```bash
npm run dev -- input.pst --output output.ics
```

### Code Style

We use ESLint and Prettier to maintain consistent code style:

```bash
# Check for linting errors
npm run lint

# Auto-fix linting issues
npm run lint:fix

# Format code with Prettier
npm run format
```

**Code Style Guidelines:**
- TypeScript with strict mode enabled
- Use meaningful variable and function names
- Add JSDoc comments for public APIs
- Follow existing patterns in the codebase
- Use 2 spaces for indentation
- Single quotes for strings
- Semicolons required

### Testing

```bash
# Run all tests
npm test

# Run tests in watch mode
npm run test:watch

# Run tests with coverage
npm run test:coverage
```

**Testing Guidelines:**
- Write tests for new features
- Maintain or improve code coverage (minimum 60%)
- Test edge cases and error conditions
- Use descriptive test names

## Making Changes

### Commit Messages

Write clear, concise commit messages:

```
feat: add support for recurring appointments
fix: correct timezone handling for all-day events
docs: update README with new CLI options
refactor: simplify calendar extraction logic
test: add tests for property mapping
```

**Format:**
- Use present tense ("add feature" not "added feature")
- Use imperative mood ("move cursor to..." not "moves cursor to...")
- Start with a type: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`
- Keep first line under 72 characters
- Add detailed description in body if needed

### Pull Request Process

1. **Update documentation** if you're changing functionality
2. **Add tests** for new features or bug fixes
3. **Ensure all tests pass**: `npm test`
4. **Ensure code is formatted**: `npm run format`
5. **Ensure linting passes**: `npm run lint`
6. **Build successfully**: `npm run build`
7. **Update CHANGELOG.md** if appropriate
8. **Push to your fork** and submit a pull request

### Pull Request Guidelines

- Fill out the pull request template completely
- Link related issues using keywords (e.g., "Fixes #123")
- Keep pull requests focused on a single feature or fix
- Add screenshots for UI changes (if applicable)
- Be responsive to feedback and questions

## Project Structure

```
fixEcalendar/
├── src/
│   ├── index-with-db.ts         # Main CLI with database support
│   ├── index.ts                 # Legacy CLI
│   ├── database/
│   │   └── calendar-db.ts       # SQLite database manager
│   ├── parser/
│   │   ├── pst-parser.ts        # PST file parser
│   │   ├── calendar-extractor.ts # Calendar extraction
│   │   └── types.ts             # Type definitions
│   ├── converter/
│   │   ├── ical-converter.ts    # iCal generator
│   │   ├── property-mapper.ts   # Property mapping
│   │   └── types.ts             # Type definitions
│   ├── utils/
│   │   ├── logger.ts            # Logging utilities
│   │   ├── validators.ts        # Input validation
│   │   └── error-handler.ts    # Error handling
│   └── config/
│       └── constants.ts         # Configuration constants
├── tests/                        # Test files
├── dist/                        # Compiled output (generated)
└── node_modules/                # Dependencies (generated)
```

## Areas for Contribution

We welcome contributions in these areas:

### High Priority
- Test coverage improvements
- Bug fixes
- Performance optimizations
- Documentation improvements

### Feature Requests
- Attachment extraction and linking
- Category/tag mapping to iCal CATEGORIES
- Enhanced timezone detection
- Email address validation
- Contact extraction (separate feature)
- Task extraction (separate feature)

### Infrastructure
- CI/CD improvements
- Docker containerization
- npm package publication preparation

## Reporting Bugs

Use the GitHub issue tracker to report bugs. When filing a bug report, include:

- fixEcalendar version (`npm list fixecalendar`)
- Node.js version (`node --version`)
- Operating system and version
- PST file format (ANSI or Unicode) if relevant
- Clear description of the issue
- Steps to reproduce
- Expected vs actual behavior
- Error messages and stack traces
- Sample PST file (if possible and doesn't contain sensitive data)

## Suggesting Enhancements

We're open to enhancement suggestions! When suggesting an enhancement:

- Check if it's already suggested in existing issues
- Provide a clear use case
- Explain why this enhancement would be useful
- Consider implementation approaches
- Be open to discussion and alternatives

## Security Issues

Please report security vulnerabilities privately. See [SECURITY.md](SECURITY.md) for details.

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

## Questions?

Feel free to open a discussion on GitHub or reach out through the issue tracker.

Thank you for contributing to fixEcalendar!
