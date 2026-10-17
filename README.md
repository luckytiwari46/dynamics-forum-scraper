# Dynamics Forum Scraper

A Python-based web scraping and browser automation project for extracting discussion content from Microsoft Dynamics Community forums and exporting it into structured Microsoft Word documents.

## Overview

Dynamics Forum Scraper automates the process of collecting forum discussions from Microsoft Dynamics Community archive pages. It uses browser automation to navigate forum pages, extract topic content, and save the results as individual Word documents.

The project supports both single-page processing and batch processing across a specified range of archive pages. Performance optimizations help reduce unnecessary browser operations and improve execution time while aiming to preserve the existing extraction logic and document formatting.

## Features

- **Automated Forum Extraction:** Collects discussion content from Microsoft Dynamics Community forum pages.
- **Page-Range Processing:** Processes a user-defined range of archive pages.
- **Individual Word Documents:** Generates a separate `.docx` file for each processed archive page.
- **Structured Content Export:** Preserves supported document elements, including headings, text, author information, dates, links, tables, code blocks, and images.
- **Browser Automation:** Uses Playwright to navigate pages and interact with dynamic web content.
- **Performance Optimization:** Reduces redundant browser operations, unnecessary waits, and repeated DOM queries.
- **Image Handling:** Supports saving extracted images alongside the generated documents.
- **Command-Line Workflow:** Accepts the archive URL and page range through the terminal.

## Technologies Used

- Python
- Playwright
- Microsoft Word `.docx` document generation
- Browser automation and JavaScript DOM evaluation
- Windows Command Prompt and batch scripting

## How It Works

1. Run the project using its batch file or Python script.
2. Enter the Microsoft Dynamics Community archive URL.
3. Specify the starting and ending page numbers.
4. The program processes each archive page in the selected range.
5. It extracts the available discussion content from the topics on each page.
6. It generates individual Word documents using the project's existing document-generation logic.
7. The output documents and associated image folders are saved to the configured output directory.

## Example

Archive URL:

`https://community.dynamics.com/forums/thread/?groupid=fca35e6b-9686-ef11-ac21-7c1e520a09df&page=872&sortorder=mspwrplt_cmty_latestactivitydate&status=&category=&posttype=&tab=public`

Starting page: `872`

Ending page: `882`

The program processes pages 872 through 882 and generates separate Word documents for each page, following the project's existing filename convention.

## Installation

1. Install Python on your computer.
2. Clone or download this repository.
3. Install the Python dependencies required by the project.
4. Install the browser binaries required by the Playwright configuration, if they are not already installed.

Check the project files for the exact Python version, dependency list, and browser setup instructions.

## Usage

Run the provided `.bat` file from Windows Command Prompt, or execute the Python script using the appropriate Python environment.

Follow the prompts to enter the archive URL and page range.

For single-page processing, enter the same number for both the starting and ending pages.

## Output

The program generates Microsoft Word `.docx` files for the processed archive pages and stores associated images in separate folders when required by the existing implementation.

Example filenames:

- `Page_0872.docx`
- `Page_0873.docx`
- `Page_0874.docx`

The actual output directory and naming convention depend on the project configuration.

## Project Goals

- Automate repetitive forum data collection.
- Make large archive ranges easier to process.
- Reduce unnecessary browser automation overhead.
- Organize extracted discussions into readable Word documents.
- Maintain the existing single-page workflow while supporting batch processing.

## Disclaimer

This project is intended for legitimate research, learning, and personal or authorized archival purposes. Users are responsible for complying with the website's terms of service, applicable copyright requirements, privacy rules, and rate limits. Avoid collecting private or sensitive information without authorization.

## Author

Lucky Tiwari

GitHub: https://github.com/luckytiwari46

LinkedIn: https://linkedin.com/in/tiwarilucky
