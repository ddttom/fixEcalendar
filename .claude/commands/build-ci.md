Simulate the GitHub Actions CI pipeline locally with the following command:

```bash
npm ci && npm run lint && npm run format -- --check && npm run build && npm run test:coverage
```

This command will execute in sequence, stopping immediately if any step fails:

1. **Clean install** - Fresh dependency install with npm ci (reproducible)
2. **Lint check** - Check for ESLint issues
3. **Format check** - Verify code formatting without modifying files
4. **Build** - Compile TypeScript to dist/
5. **Test with coverage** - Run full test suite and generate coverage report

If any command fails, stop execution and display the error. Do not continue to subsequent steps.

This replicates what runs in CI/CD and helps catch failures before pushing to remote.
