Run a test-focused build workflow with the following command:

```bash
npm run build && npm test
```

This command will execute in sequence, stopping immediately if any step fails:

1. **Build** - Compile TypeScript to dist/
2. **Run tests** - Execute full Jest test suite

If any command fails, stop execution and display the error. Do not continue to subsequent steps.

This is useful for quick iteration on test development, ensuring code compiles before running tests.
