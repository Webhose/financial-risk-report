# Financial Risk Report Generator

## Overview
This Python script automates the generation of weekly financial risk reports. It utilizes Webz.io to fetch news articles with a negative sentiment in the domains of economy, business, and finance. These articles are then analyzed using ChatGPT to identify financial risks, forming the basis of structured financial risk reports. Additionally, the script generates a featured image for each report using OpenAI's DALL-E model.

## Requirements
- Python 3.x
- `docx`: For creating Word documents
- `requests`: To make HTTP requests
- `openai`: OpenAI's Python client library
- `python-Levenshtein`: For computing string similarities
- `beautifulsoup4`: For HTML parsing

## Setup
1. Install the necessary Python packages:
   ```bash
   pip install python-docx requests openai python-Levenshtein beautifulsoup4
   ```
2. Ensure the environment variables `WEBZ_API_KEY` and `OPENAI_API_KEY` are set with your respective Webz.io and OpenAI API keys.

## Usage
To run the script, use:
```bash
python create_financial_risk_report.py
```
### Script Workflow
- Fetches negative sentiment news articles related to economy, business, and finance from Webz.io.
- Filters out similar articles using the Levenshtein ratio.
- Utilizes ChatGPT to generate a detailed financial risk analysis report for each article.
- Creates a Word document containing these reports, complete with an auto-generated title and introduction.
- Generates and adds a featured image for each report using DALL-E.

## Functions Description
- `fetch_articles(query, api_key, total)`: Fetches news articles from Webz.io's API.
- `generate_reports(filtered_articles)`: Produces financial risk reports from the filtered articles.
- `create_word_doc(file_name, title_text, image_url, intro, reports)`: Generates a Word document with the reports and additional information.
- Various utility functions for text processing and HTML content handling.

## Output
Produces a Word document titled "financial risk digest.docx", encompassing the financial risk reports, an introduction, a title, and a DALL-E generated image.

## Notes
- Internet access is required for API calls and image generation.
- The quality and accuracy of the reports are contingent on the input articles and the performance of the ChatGPT model.
