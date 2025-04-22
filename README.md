# Look for FTEs and Linkedin

This script is designed to enrich a company dataset by retrieving employee count estimates and official LinkedIn URLs for companies using web scraping and the Google Search API.

## Features

- Searches Google for each company's LinkedIn page.
- Extracts estimated employee counts from the LinkedIn page.
- Updates missing fields (`FTE`, `Company linkedin `) in an Excel file.
- Handles errors gracefully with exponential backoff for web requests.

## Requirements

Make sure you have the following Python libraries installed:

```bash
pip install pandas googlesearch-python requests openpyxl
