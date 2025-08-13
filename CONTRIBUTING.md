# Contributing to Universal Document Processor

Thank you for your interest in contributing to Universal Document Processor! This document provides guidelines and instructions for contributing to this project.

## Code of Conduct

By participating in this project, you agree to abide by our [Code of Conduct](CODE_OF_CONDUCT.md). Please read it before contributing.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check the issue list to avoid duplicates. When creating a bug report, include as many details as possible:

- Use a clear and descriptive title
- Describe the exact steps to reproduce the problem
- Describe the behavior you observed and what you expected to see
- Include screenshots if possible
- Include details about your environment (OS, Node.js version, etc.)
- Use the bug report template when creating an issue

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion:

- Use a clear and descriptive title
- Provide a detailed description of the proposed enhancement
- Explain why this enhancement would be useful
- Include examples of how it would be used
- Use the feature request template when creating an issue

### Pull Requests

- Fill in the required template
- Follow the coding style and conventions
- Include appropriate tests
- Update documentation as needed
- Ensure all tests pass
- Link to any related issues

#### Intellectual Property License

**By submitting a Pull Request, you agree that all contributions submitted through Pull Requests will be licensed under the MIT License.** This means:

- You grant the project maintainers and users the right to use, modify, and distribute your contributions under the MIT License
- You confirm that you have the right to make these contributions
- You understand that your contributions will become part of the open source project
- You waive any claims to exclusive ownership of the contributed code

If you cannot agree to these terms, please do not submit a Pull Request.

## Development Setup

1. Fork and clone the repository
2. Install dependencies:
   ```bash
   pnpm install
   ```
3. Set up environment variables:
   ```bash
   cp env.example .env
   ```
4. Run tests to ensure everything is working:
   ```bash
   pnpm test
   ```

## Development Workflow

1. Create a new branch for your feature or bugfix:
   ```bash
   git checkout -b feature/your-feature-name
   ```
   or
   ```bash
   git checkout -b fix/your-bugfix-name
   ```

2. Make your changes and commit them with a descriptive message:
   ```bash
   git commit -m "Add feature: your feature description"
   ```

3. Push your branch to your fork:
   ```bash
   git push origin feature/your-feature-name
   ```

4. Create a Pull Request from your fork to the main repository

## Coding Guidelines

### JavaScript/TypeScript

- Use TypeScript for new code
- Follow the ESLint and Prettier configurations
- Write JSDoc comments for functions and classes
- Use async/await for asynchronous code
- Handle errors appropriately

### Testing

- Write tests for new features and bug fixes
- Ensure all tests pass before submitting a PR
- Aim for good test coverage

### Documentation

- Update documentation for new features or changes
- Use clear and concise language
- Include examples where appropriate

## Project Structure

- `src/`: Source code
  - `index.ts`: Main entry point
  - `tools/`: Core functionality modules
  - `types/`: TypeScript type definitions
  - `utils/`: Utility functions
- `tests/`: Test files
  - `unit/`: Unit tests
  - `integration/`: Integration tests
  - `fixtures/`: Test fixtures
- `docs/`: Documentation
  - `api/`: API documentation
  - `examples/`: Usage examples
- `examples/`: Example code
- `scripts/`: Build and development scripts

## Release Process

1. Update version in package.json
2. Update CHANGELOG.md
3. Create a new tag with the version number
4. Push the tag to trigger the release workflow

## Questions?

If you have any questions, feel free to create an issue or contact the maintainer at teleaistudio@gmail.com.