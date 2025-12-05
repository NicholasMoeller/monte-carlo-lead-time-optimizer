# Changelog

All notable changes to the Shared Resource Monte Carlo Simulation project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2025-12-05

### Added
- **Configuration Constants**: Centralized constants at top of VBA module for easy customization
  - `MAIN_SIMULATION_COUNT` - Main simulation iterations (default: 2000)
  - `VOLATILITY_CHART_SIMS` - Volatility chart simulations (default: 100)
  - `SIMULATION_DAYS` - Forecast horizon in days (default: 365)
  - `RISK_PERCENTILE` - Risk level for percentile calculation (default: 0.95)
  - `VOLATILITY_THRESHOLD` - CV threshold for volatile products (default: 0.3)
  - `OVERLOAD_BUFFER` - Capacity buffer for overload detection (default: 1.1)
- **Enhanced Documentation**: Comprehensive function headers with parameters, algorithms, and outputs
- **Code Optimization**: Refactored hardcoded values to use configuration constants
- **Repository Files**: Added LICENSE (MIT), CHANGELOG.md, and .gitignore
- **Version Information**: Added VERSION constant for tracking releases

### Changed
- Increased volatility chart simulations from 50 to 100 for denser visualization
- Updated chart title and legend to reflect 100 simulation runs
- Improved code organization with clear sections and documentation

### Fixed
- Auto-sizing for columns and rows to prevent text smooshing
- Overload message formatting for better readability

## [1.0.0] - 2025-11-29

### Added
- **Core Monte Carlo Simulation**: 2,000-iteration simulation for safe lead time calculation
- **Shared Resource Modeling**: "Shared fate" algorithm for products on same production line
- **Statistical Analysis**: Automatic calculation of average demand and standard deviation from sales history
- **Dual Buffer Methodologies**:
  - Equal Treatment (Column G): Uniform buffer allocation
  - Risk-Based (Column H): Variance-weighted buffer allocation
- **System Buffer Analysis** (Columns J & K):
  - Toxic Threshold: Spare capacity indicator
  - Max Safe Order Qty: Maximum order before overload
  - Color-coded health indicators (Green/Orange/Red)
- **Large Order Detection** (Columns L & M):
  - Large Order Quantity: Statistical threshold (Avg + 2σ)
  - Large Order LT Quote: Extended lead time for large orders
- **Volatility Spaghetti Chart**: Visual proof of demand uncertainty with 50 simulation runs
- **Material Lead Time Support** (Column I): Product-specific material procurement time
- **Overload Detection**: Automatic detection and warning for overloaded production lines
- **Performance Optimizations**: Screen updating and calculation disabled during processing
- **Error Handling**: Comprehensive validation and user-friendly error messages
- **Visual Feedback**: Color-coded results (Green/Orange/Red/Pink/Yellow)

### Features
- Groups products by production line automatically
- Handles multiple production lines in one run
- Validates worksheet structure before processing
- Displays execution time on completion
- Auto-resizes columns and rows for clean display

### Documentation
- Comprehensive README with formulas, examples, and troubleshooting
- Sample CSV template with all column headers
- Detailed installation and usage instructions
- Version history and customization guide

---

## Version Numbering

This project follows [Semantic Versioning](https://semver.org/):
- **MAJOR** version for incompatible API/structure changes
- **MINOR** version for new functionality in a backward-compatible manner
- **PATCH** version for backward-compatible bug fixes

---

## Categories

- **Added**: New features
- **Changed**: Changes in existing functionality
- **Deprecated**: Soon-to-be removed features
- **Removed**: Removed features
- **Fixed**: Bug fixes
- **Security**: Vulnerability fixes
