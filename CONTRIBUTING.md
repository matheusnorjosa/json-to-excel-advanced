# Contributing to json-to-excel-advanced

First off, thank you for considering contributing to json-to-excel-advanced! 🎉

## Code of Conduct

This project and everyone participating in it is governed by respect and professionalism. Please be kind and courteous to other contributors.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check the existing issues to avoid duplicates. When you create a bug report, include as many details as possible:

- **Use a clear and descriptive title**
- **Describe the exact steps to reproduce the problem**
- **Provide specific examples** (sample JSON files are very helpful)
- **Describe the behavior you observed and what you expected**
- **Include your Python version** and dependency versions

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion, include:

- **Use a clear and descriptive title**
- **Provide a detailed description of the suggested enhancement**
- **Explain why this enhancement would be useful**
- **List any alternative solutions you've considered**

### Pull Requests

1. Fork the repo and create your branch from `main`
2. If you've added code, add tests
3. Ensure the test suite passes
4. Make sure your code follows the existing style
5. Write a clear commit message

## Development Setup

```bash
# Clone your fork
git clone https://github.com/YOUR_USERNAME/json-to-excel-advanced.git
cd json-to-excel-advanced

# Install dependencies
pip install -r requirements.txt

# Run the example
python json_to_excel.py example.json
```

## Testing

Before submitting a pull request, make sure to test your changes:

```bash
# Test with the example file
python json_to_excel.py example.json -o test_output.xlsx

# Test with your own JSON files
python json_to_excel.py your_file.json
```

## Style Guidelines

### Python Style

- Follow PEP 8
- Use meaningful variable names
- Add docstrings to functions and classes
- Keep functions focused and small
- Add comments for complex logic

### Commit Messages

- Use the present tense ("Add feature" not "Added feature")
- Use the imperative mood ("Move cursor to..." not "Moves cursor to...")
- Limit the first line to 72 characters or less
- Reference issues and pull requests after the first line

Example:
```
Add support for custom date formats

- Implement date_format configuration option
- Add tests for different date formats
- Update documentation

Fixes #123
```

## Project Structure

```
json-to-excel-advanced/
├── json_to_excel.py      # Main converter script
├── README.md             # Project documentation
├── LICENSE               # MIT License
├── requirements.txt      # Python dependencies
├── .gitignore           # Git ignore file
├── example.json         # Example input file
└── CONTRIBUTING.md      # This file
```

## Adding New Features

When adding a new feature:

1. **Discuss first** - Open an issue to discuss your idea
2. **Keep it simple** - Maintain the tool's ease of use
3. **Document it** - Update README.md with usage examples
4. **Test it** - Ensure it works with various JSON structures

## Questions?

Feel free to open an issue with the "question" label, or reach out through GitHub Discussions.

Thank you for contributing! 🚀
