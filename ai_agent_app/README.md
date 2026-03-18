# AI Agent - Office Automation

A desktop automation app for Word, Excel, PowerPoint, PDF, OCR, and basic browser workflows. It includes a PySide6 UI, AI chat commands, direct quick actions, and local file automation on Windows.

---

## Overview

This app can:

- create and edit Word documents
- open and control live Excel workbooks
- create and update PowerPoint presentations
- extract text and images from PDFs
- run OCR on images and save the result to TXT or Word
- open browsers, close them, and scrape webpage content
- use either OpenAI or Gemini-based tool-calling agents depending on the app setup

---

## Project Structure

```text
ai_agent_app/
|-- main.py
|-- requirements.txt
|-- setup.bat
|-- src/
|   |-- agents/
|   |   |-- gemini_agent.py
|   |   `-- openai_agent.py
|   |-- tools/
|   |   |-- browser_tools.py
|   |   |-- excel_tools.py
|   |   |-- executor.py
|   |   |-- ocr_tools.py
|   |   |-- pdf_tools.py
|   |   |-- ppt_tools.py
|   |   `-- word_tools.py
|   `-- ui/
|       `-- main_window.py
```

---

## Setup

### Windows quick setup

```bat
setup.bat
```

### Manual setup

```bash
pip install -r requirements.txt
```

### Optional browser rendering setup

For JavaScript-rendered scraping and interactive browser control:

```bash
pip install playwright
playwright install
```

---

## Run

```bash
cd ai_agent_app
python main.py
```

---

## API Keys

### OpenAI

Add your key in the sidebar or in `.env`:

```env
OPENAI_API_KEY=your_key_here
```

### Gemini

If you use the Gemini flow, create a key from:

`https://aistudio.google.com/apikey`

---

## Features

### Word

- Create `.docx` files with title, font size, and color
- Add paragraphs and headings
- Format matching text with bold, italic, color, and size
- Open created Word files automatically

### Excel

- Open the Excel desktop app
- Create or open workbooks and sheets
- Write and read ranges in the active workbook
- Apply formatting such as bold, fill color, font color, number format, alignment, and autofit
- Save and close workbooks
- Create file-based demo or table workbooks

### PowerPoint

- Create and open presentations automatically
- Create title slides with title, subtitle, company, and date
- Add content slides with bullet points
- Control font size, text color, fill/background color, and font family such as `Aptos`
- Edit slides
- Add slides even when the presentation is already open

### PDF

- Extract images from PDFs
- Extract text from PDFs
- Convert Word to PDF

### OCR

- Extract text from images
- Save OCR output to `.txt`
- Save OCR output directly to Word

### Browser and Web Scraping

- Open Edge, Chrome, or Firefox
- Close the last opened browser or all windows for a browser
- Scrape static webpage text and links
- Scrape JavaScript-rendered pages with Playwright
- Open an interactive browser session, manually inspect or verify a page, then scrape the current page

---

## Example Prompts

### Word

```text
Create a Word file at C:/reports/q4.docx with title Q4 Report
Open C:/reports/q4.docx and make the word Revenue bold and blue
Add a heading Summary to C:/reports/q4.docx
```

### Excel

```text
Open Excel and create a March sales sheet
Write data starting at A1 in the active Excel workbook
Read cells A1:D10 only from the active Excel workbook
Format range A1:D10 in the active Excel workbook with bold headers and autofit
```

### PowerPoint

```text
Create and open a PowerPoint titled Company Overview
Create and open a PowerPoint about Data Science with title text color red and title font size 8
Create and open a PowerPoint about Data Science with title font Aptos
Add a slide titled Introduction to the open PowerPoint
```

### PDF

```text
Extract images from C:/docs/report.pdf to C:/output/images
Extract text from page 2 of C:/docs/report.pdf
Convert C:/reports/q4.docx to PDF at C:/reports/q4.pdf
```

### OCR

```text
Extract text from image C:/images/note.png
Save OCR text from image C:/images/note.png to Word file C:/output/note.docx
```

### Browser

```text
Open Edge and go to google.com
Scrape https://example.com and return up to 500 characters
Scrape https://www.python.org with JavaScript rendering using Edge and return up to 1000 characters
Open Edge interactively and go to https://openai.com
Scrape the current browser page
Close the browser
```

---

## Browser Tools Help

### When to use each browser command

- Use `Open Edge and go to ...` when you just want to launch a browser tab
- Use `Scrape https://...` for normal static webpage scraping
- Use `... with JavaScript rendering using Edge` for modern JS-heavy pages
- Use `Open Edge interactively and go to ...` when you want to manually inspect, log in, or complete verification before scraping
- Use `Scrape the current browser page` only after opening an interactive browser session
- Use `Close the browser` to close the tracked browser session

### Recommended browser flow

#### Static page flow

```text
Open Edge and go to https://example.com
Scrape https://example.com and return up to 1000 characters
Close the last opened browser
```

#### JavaScript-rendered flow

```text
Scrape https://www.python.org with JavaScript rendering using Edge and return up to 1000 characters
```

#### Interactive manual flow

```text
Open Edge interactively and go to https://openai.com
Scrape the current browser page
Close the browser
```

In the interactive flow, you can manually log in, solve a verification step, or navigate to another page before running `Scrape the current browser page`.

### Good test prompts

```text
Open Edge and go to https://www.python.org
Scrape https://example.com and include links
Scrape https://www.python.org with JavaScript rendering using Edge and return up to 1000 characters
Open Edge interactively and go to https://example.com
Scrape the current browser page
```

### Good websites for testing

- `https://example.com`
- `https://example.org`
- `https://httpbin.org/html`
- `https://www.python.org`

### Limitations

- Some websites may return anti-bot, verification, or placeholder pages
- JavaScript-rendered scraping works best when Playwright is installed correctly
- Interactive browser scraping currently works best with Edge and Chrome
- Browser content may differ from what you see manually on highly protected websites

---

## Notes

- Use Windows-style absolute paths such as `C:/output/file.docx`
- Output folders are created automatically when possible
- Browser-rendered scraping may still hit anti-bot or verification pages on some websites
- Interactive browser scraping works best after you manually complete login or verification steps
- This app performs file operations locally; only AI requests go to the configured model provider
