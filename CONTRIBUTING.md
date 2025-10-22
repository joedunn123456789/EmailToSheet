# Contributing to Email to Excel Exporters

Thank you for considering contributing! 🎉

## How to Contribute

### Reporting Bugs 🐛

If you find a bug, please [open an issue](../../issues) with:
- A clear title
- Description of the problem
- Steps to reproduce
- Expected vs actual behavior
- Your environment (OS, Python version, etc.)

### Suggesting Features 💡

Have an idea? [Open an issue](../../issues) with:
- Clear description of the feature
- Why it would be useful
- Example use cases

### Submitting Pull Requests 🔀

1. Fork the repository
2. Create a new branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Code Guidelines

**Python Code:**
- Follow PEP 8 style guide
- Add comments explaining complex logic
- Include docstrings for functions
- Test with multiple email accounts if possible

**Google Apps Script:**
- Use clear variable names
- Add comments explaining what each section does
- Follow existing code style

**Documentation:**
- Keep README updated
- Add examples for new features
- Use clear, simple language
- Include "Explain It Like You're 5" sections where helpful

## Development Setup

### For Python Scripts
```bash
# Clone the repo
git clone https://github.com/yourusername/email-to-excel.git
cd email-to-excel

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install openpyxl

# Run tests (if you add them)
python -m pytest
```

### For Google Apps Script
- Open Google Sheets
- Extensions → Apps Script
- Copy script code for testing

## Ideas for Contributions

- [ ] Add support for Yahoo Mail
- [ ] Add attachment downloading
- [ ] Create GUI version
- [ ] Add email filtering by date range
- [ ] Add progress bar for large exports
- [ ] Add email deduplication
- [ ] Create Docker container version
- [ ] Add unit tests
- [ ] Add support for multiple folders at once
- [ ] Create web interface
- [ ] Add HTML email formatting preservation
- [ ] Add email threading/conversation grouping

## Code of Conduct

- Be respectful and considerate
- Welcome newcomers
- Focus on what's best for the community
- Show empathy towards others

## Questions?

Feel free to [open an issue](../../issues) with the "question" label!

Thank you for contributing! 🙏
