# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based web automation tool for Microsoft Dynamics CRM. It automates the bulk updating of knowledge base articles by reading data from an Excel file and using Selenium WebDriver to interact with the CRM web interface.

## Commands

### Running the Application
```bash
# Activate virtual environment (Windows)
venv\Scripts\activate

# Run the main application
python -m app.main
```

### Installing Dependencies
```bash
# Install required packages
pip install -r requirements.txt
```

### Required Setup
- Chrome browser must be installed
- ChromeDriver must be installed and in PATH
- Excel data file must be present at `files/data.xlsx` with columns:
  - '記事' (Article GUID)
  - '番号' (KBA number)
  - '対象' (Target audience)

## Architecture

The application follows a simple layered architecture:

```
main.py                 # Entry point - orchestrates the update process
├── config.py          # Configuration settings and logging setup
├── controller.py      # Business logic - reads Excel data and coordinates updates
└── dynamics.py        # CRM interaction layer using Selenium WebDriver
```

Key architectural decisions:
- **Selenium-based automation**: Uses Chrome WebDriver to interact with Dynamics CRM web interface
- **Retry mechanism**: Built-in retry logic (3 attempts) for handling transient failures
- **Error resilience**: Continues processing remaining articles even if individual updates fail
- **Logging**: Comprehensive logging to both file (`log/app.log`) and console

## Important Implementation Details

### CRM Interaction Flow (dynamics.py)
1. Navigate to article URL using GUID
2. Unpublish article if published
3. Update title with "【メンテ済】" prefix
4. Update target audience based on Excel data
5. Save and republish article

### Error Handling
- Custom `PublishButtonNotActiveError` for specific UI state issues
- Retry logic for all CRM operations
- Detailed error logging with Japanese messages
- Summary reporting of successes/failures

### Configuration (config.py)
- HEADLESS_MODE: False (Chrome runs with UI)
- RETRY_COUNT: 3
- SELENIUM_WAIT_TIME: 5 seconds
- PAGE_LOAD_WAIT: 3 seconds

## Development Notes

- The application uses Japanese for all user-facing messages and logs
- Internal CRM URL: `http://sv-vw-ejap:5555/main.aspx?app=d365default&pagetype=entityrecord&etn=knowledgearticle&id={article_id}`
- Windows authentication may be required for CRM access
- No unit tests are currently implemented