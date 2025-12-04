Run a quick build verification with the following command:

```bash
npm run lint && npm run build
```

This command will execute in sequence, stopping immediately if any step fails:

1. **Lint check** - Check for ESLint issues (no auto-fix)
2. **Build** - Compile TypeScript to dist/

If any command fails, stop execution and display the error. Do not continue to subsequent steps.

This is a faster feedback loop for development, skipping tests and formatting.
