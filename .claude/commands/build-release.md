Run comprehensive pre-release verification with the following command:

```bash
rm -rf dist/ coverage/ && npm ci && npm run lint && npm run test:coverage && npm run build && ls -la dist/ && cat package.json | grep '"version"'
```

This command will execute in sequence, stopping immediately if any step fails:

1. **Clean artifacts** - Remove dist/ and coverage/ directories
2. **Clean install** - Fresh dependency install with npm ci
3. **Lint check** - Check for ESLint issues
4. **Test with coverage** - Run full test suite and generate coverage report
5. **Build** - Compile TypeScript to dist/
6. **Verify artifacts** - List dist/ contents to confirm build output
7. **Display version** - Show current version from package.json

If any command fails, stop execution and display the error. Do not continue to subsequent steps.

This ensures the project is in a release-ready state with full verification from a clean slate.
