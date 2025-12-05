# Contributing to Shared Resource Monte Carlo Simulation

Thank you for your interest in contributing! This document provides guidelines and instructions for contributing to this project.

## 📋 Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How Can I Contribute?](#how-can-i-contribute)
- [Development Setup](#development-setup)
- [Coding Standards](#coding-standards)
- [Testing Guidelines](#testing-guidelines)
- [Submitting Changes](#submitting-changes)
- [Documentation](#documentation)

---

## 🤝 Code of Conduct

### Our Pledge

We are committed to providing a welcoming and inspiring community for all. Please:

- Use welcoming and inclusive language
- Be respectful of differing viewpoints and experiences
- Accept constructive criticism gracefully
- Focus on what is best for the community
- Show empathy towards other community members

### Unacceptable Behavior

- Harassment, trolling, or derogatory comments
- Public or private intimidation
- Publishing others' private information without permission
- Any conduct that could be considered inappropriate in a professional setting

---

## 🎯 How Can I Contribute?

### Reporting Bugs

Before creating a bug report:
1. Check the [troubleshooting guide](README.md#troubleshooting)
2. Search [existing issues](../../issues) to avoid duplicates
3. Test with the latest version

When reporting a bug, include:
- **Description**: Clear summary of the issue
- **Steps to Reproduce**: Numbered list of exact steps
- **Expected Behavior**: What should happen
- **Actual Behavior**: What actually happens
- **Environment**: Excel version, OS, VBA version
- **Sample Data**: Anonymized data if possible (CSV format)
- **Screenshots**: If applicable

**Example Bug Report**:
```markdown
### Bug: Overload detection incorrect for capacity = 0

**Steps to Reproduce**:
1. Set Line Capacity to 0 in Simulation sheet
2. Run the macro

**Expected**: Error message about invalid capacity
**Actual**: Divide by zero error

**Environment**: Excel 2019, Windows 10
```

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. Include:

- **Use Case**: Why is this enhancement needed?
- **Proposed Solution**: How should it work?
- **Alternatives**: Other approaches considered
- **Examples**: Sample scenarios demonstrating the value
- **Backward Compatibility**: Impact on existing users

**Example Enhancement**:
```markdown
### Enhancement: Add support for weekly capacity constraints

**Use Case**: Some production lines only operate 5 days/week

**Proposed Solution**: Add "Days Per Week" column to adjust capacity calculations

**Backward Compatibility**: Default to 7 days/week for existing users
```

### Pull Requests

We actively welcome pull requests for:

- Bug fixes
- New features
- Documentation improvements
- Performance optimizations
- Test coverage improvements
- Code refactoring

---

## 🛠️ Development Setup

### Prerequisites

- **Microsoft Excel** (2013 or later)
- **VBA Editor** enabled (Developer tab in Excel)
- **Git** for version control

### Setting Up Your Development Environment

1. **Fork the Repository**
   ```bash
   # On GitHub, click "Fork" button
   ```

2. **Clone Your Fork**
   ```bash
   git clone https://github.com/YOUR_USERNAME/Work.git
   cd Work
   ```

3. **Create a Branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

4. **Import VBA Module**
   - Open Excel
   - Press `Alt+F11` to open VBA Editor
   - File → Import → Select `SharedResourceMonteCarloSimulation.bas`

5. **Set Up Test Data**
   - Create "Simulation" and "SalesHistory" sheets
   - Use `sample_data/Simulation.csv` as template
   - Add your own test cases

---

## 📝 Coding Standards

### VBA Conventions

#### Naming
```vba
' Constants: UPPER_CASE with underscores
Private Const MAX_ITERATIONS As Long = 2000

' Functions/Subs: PascalCase
Public Sub CalculateStatistics()

' Variables: camelCase
Dim productCount As Long
Dim totalAvgDemand As Double

' Parameters: camelCase
Private Sub ProcessLine(lineName As String, capacity As Double)
```

#### Comments
```vba
' ==================================================================================
' Function Name - Brief Description
' ==================================================================================
' Description:  Detailed explanation of purpose
'
' Parameters:
'   param1  - Description of parameter 1
'   param2  - Description of parameter 2
'
' Returns:      Description of return value (for functions)
'
' Algorithm:
'   1. Step one
'   2. Step two
'
' Error Handling: How errors are handled
' ==================================================================================
Public Function MyFunction(param1 As String) As Double
    ' Implementation
End Function
```

#### Error Handling
```vba
' Always include error handling for public functions
Public Sub MyFunction()
    On Error GoTo ErrorHandler

    ' Your code here

CleanUp:
    ' Cleanup code (restore settings, etc.)
    Exit Sub

ErrorHandler:
    ' Handle errors
    MsgBox "Error in MyFunction: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
```

#### Code Organization
```vba
' 1. Module-level constants
Private Const CONSTANT_NAME As Type = Value

' 2. Module-level variables (minimize these)
Private moduleVariable As Type

' 3. Public functions/subs
Public Sub PublicFunction()
End Sub

' 4. Private helper functions
Private Sub HelperFunction()
End Sub
```

### Style Guidelines

- **Indentation**: 4 spaces (VBA default)
- **Line Length**: Max 100 characters (use line continuation `_` if needed)
- **Option Explicit**: Always use at module top
- **Variable Declarations**: Declare all variables explicitly
- **Magic Numbers**: Use named constants instead of hardcoded values

**Good Example**:
```vba
Private Const PERCENTILE_95 As Double = 0.95
Dim result As Double
result = Percentile(data, PERCENTILE_95)
```

**Bad Example**:
```vba
Dim result
result = Percentile(data, 0.95)  ' Magic number, no type declaration
```

---

## 🧪 Testing Guidelines

### Manual Testing

Before submitting a PR, test with:

1. **Typical Case**: Standard 3-product scenario
2. **Edge Cases**:
   - Single product per line
   - 10+ products per line
   - Capacity = 0 (should error)
   - Negative demand values
   - Missing sales history
3. **Overload Scenario**: Total demand > capacity
4. **Multiple Lines**: 2-3 different production lines
5. **Large Dataset**: 50+ products across multiple lines

### Test Checklist

- [ ] Macro completes without errors
- [ ] Results appear in correct columns
- [ ] Color coding works (Green/Orange/Red)
- [ ] Volatility chart generates correctly
- [ ] Overload detection works
- [ ] Error messages are clear and helpful
- [ ] Execution time is reasonable (<10 seconds for typical data)
- [ ] No Excel crashes or hangs

### Documenting Tests

Include test results in your PR:

```markdown
## Test Results

**Environment**: Excel 2019, Windows 10

**Test Case 1: Standard 3-product line**
- ✅ Calculations correct
- ✅ Chart generated
- ✅ Execution time: 3.2 seconds

**Test Case 2: Overloaded capacity**
- ✅ Red warning displayed
- ✅ Suggested capacity shown
```

---

## 📤 Submitting Changes

### Commit Message Format

Use clear, descriptive commit messages:

```
<type>: <subject>

<body>

<footer>
```

**Types**:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation only
- `style`: Code formatting (no logic change)
- `refactor`: Code refactoring
- `test`: Adding tests
- `chore`: Maintenance tasks

**Examples**:
```
feat: Add support for custom percentile risk levels

Allow users to configure risk percentile via constant.
Default remains 95th percentile for backward compatibility.

Closes #42
```

```
fix: Prevent divide by zero when capacity is 0

Add validation check before capacity calculations.
Display user-friendly error message.

Fixes #38
```

### Pull Request Process

1. **Update Documentation**
   - Update README.md if user-facing changes
   - Update CHANGELOG.md with your changes
   - Add/update function headers if needed

2. **Update Version**
   - Increment version in VBA header
   - Follow [Semantic Versioning](https://semver.org/)
   - MAJOR.MINOR.PATCH

3. **Create Pull Request**
   - Use descriptive title
   - Reference related issues
   - Describe changes in detail
   - Include test results
   - Add screenshots if UI changes

**PR Template**:
```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Related Issues
Closes #XX

## Changes Made
- Change 1
- Change 2

## Testing
Describe testing performed

## Checklist
- [ ] Code follows style guidelines
- [ ] Self-reviewed the code
- [ ] Commented complex sections
- [ ] Updated documentation
- [ ] No new warnings
- [ ] Added tests (if applicable)
- [ ] All tests pass
```

4. **Code Review**
   - Address review feedback promptly
   - Make requested changes
   - Update PR description if scope changes
   - Be open to suggestions

5. **After Merge**
   - Delete your feature branch
   - Pull latest from main
   - Celebrate! 🎉

---

## 📚 Documentation

### When to Update Documentation

Update documentation when:
- Adding new features
- Changing behavior
- Fixing bugs that affect usage
- Adding configuration options
- Improving error messages

### Documentation Locations

- **README.md**: User-facing documentation
- **CHANGELOG.md**: Version history
- **Code Comments**: Function headers, complex logic
- **CONTRIBUTING.md**: This file

### Writing Good Documentation

**Good Example**:
```markdown
### System Buffer Health Indicator

The System Buffer (Column J) shows spare capacity:
- **Green (≥2)**: Healthy - can absorb demand spikes
- **Orange (<2)**: Fragile - vulnerable to spikes
- **Red (≤0)**: Critical - line is overloaded

**Formula**: `Capacity - Total Average Demand`
```

**Bad Example**:
```markdown
Column J shows the buffer. Green is good, red is bad.
```

---

## ❓ Questions?

- **General Questions**: Open a GitHub Discussion
- **Bug Reports**: Create an issue with `bug` label
- **Feature Requests**: Create an issue with `enhancement` label
- **Security Issues**: Email maintainers directly (see README)

---

## 🙏 Thank You!

Your contributions make this project better for everyone. We appreciate your time and effort!

**Happy coding!** 🚀
