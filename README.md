# YFinance Toolkit

A comprehensive desktop application for financial data analysis, featuring beta calculation, volatility analysis, and price data extraction from Yahoo Finance.

## 🚀 Features

- **Company Search**: Intelligent search functionality to find companies and their tickers
- **Beta Calculation**: Calculate beta coefficients against BSE Sensex benchmark
- **Volatility Analysis**: Compute annualized volatility (sigma) for selected securities
- **Price Data Export**: Extract historical price and volume data
- **Excel Reports**: Generate professionally formatted Excel reports with calculations
- **Multi-Company Analysis**: Analyze multiple securities simultaneously
- **Custom Date Ranges**: Flexible date range selection for analysis periods

## 📊 Analysis Types

### Beta Calculation
- Calculates beta coefficient relative to S&P BSE Sensex
- Uses SLOPE function for regression analysis
- Includes percentage change calculations
- Handles missing data with intelligent fill-forward

### Volatility Analysis
- Computes annualized volatility using 252 trading days
- Standard deviation-based calculation
- Professional formatting with sigma notation

### Price Data Export
- Historical close prices
- Trading volume data
- Customizable date ranges
- Clean, organized Excel output

## 🖥️ System Requirements

- Windows 10/11 (64-bit)
- Internet connection (for data fetching)
- Minimum 4GB RAM
- 100MB free disk space

## 📦 Installation

## 🔧 Dependencies

- `yfinance` - Yahoo Finance data fetching
- `pandas` - Data manipulation
- `numpy` - Numerical computations
- `openpyxl` - Excel file handling
- `customtkinter` - Modern GUI framework
- `selenium` - Web automation for enhanced search
- `beautifulsoup4` - HTML parsing
- `Google Chrome (The Browser)` - For Browsing Using Chromedriver

### Install dependencies
```bash
py -m pip install yfinance pandas numpy openpyxl customtkinter selenium beautifulsoup4 webdriver-manager
```

### Install Google Chrome 
```bash
https://www.google.com/chrome/
```

### Run from Source
```bash
# Clone the repository
git clone https://github.com/yourusername/yfinance-toolkit.git
cd yfinance-toolkit

# Install dependencies
pip install -r requirements.txt

# Run the application
python __main__.py
```

## 📖 Usage Guide

### Getting Started
1. Launch the application
2. Use the search bar to find companies
3. Click the "+" button to add companies to your selection
4. Choose your analysis type (Beta, Volatility, or Prices)
5. Set your date range
6. Click "Calculate" to generate your report

### Search Functionality
- Type partial company names or ticker symbols
- Real-time suggestions appear as you type
- Click "+" to add companies to your analysis list
- Remove companies by clicking "❌" on selected items

### Date Range Selection
- Use YYYY-MM-DD format (e.g., 2022-01-01)
- Start date must be before end date
- End date cannot be in the future
- Recommended: Use at least 1 year of data for reliable calculations

### Output Files
Reports are saved in the application directory:
- `BetaReport.xlsx` - Beta analysis results
- `VolatilityReport.xlsx` - Volatility calculations
- `[TickerNames].xlsx` - Price data exports

## 🏗️ Architecture

### Core Modules

- **`__main__.py`** - Main GUI application using CustomTkinter
- **`BetaCalculatorModule.py`** - Financial calculations and Excel report generation
- **`YFinanceModule2.py`** - Web scraping and company search functionality

### Key Features

- **Thread-safe Operations**: Background data fetching doesn't freeze the UI
- **Error Handling**: Robust error handling for network issues and missing data
- **Professional Reports**: Excel outputs with formatting, borders, and calculations
- **Cross-platform Paths**: Handles both development and PyInstaller bundled environments

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup
1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Commit your changes: `git commit -am 'Add feature'`
4. Push to the branch: `git push origin feature-name`
5. Submit a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This software is for educational and research purposes only. Financial data is provided "as is" without warranty. Always verify calculations and consult with financial professionals before making investment decisions.

## 🐛 Known Issues

- Chrome WebDriver may require manual updates for newer Chrome versions
- Some international exchanges may have limited data availability
- Large date ranges may take longer to process

## 📞 Support

- Create an [Issue](../../issues) for bug reports
- Check [Discussions](../../discussions) for general questions
- Review existing issues before creating new ones

## 🔄 Version History

See [CHANGELOG.md](CHANGELOG.md) for detailed version history.

---

**Built with ❤️ for the financial analysis community**
