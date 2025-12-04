Run the complete pre-commit build cycle with the following command:

```bash
npm run format && npm run lint:fix && npm test && npm run build
```

This command will execute in sequence, stopping immediately if any step fails:

1. **Format code** - Run Prettier to format all TypeScript files
2. **Fix lint issues** - Auto-fix ESLint issues
3. **Run tests** - Execute full Jest test suite
4. **Build** - Compile TypeScript to dist/

If any command fails, stop execution and display the error. Do not continue to subsequent steps.

On success, all code quality checks pass and the project is ready to commit.
