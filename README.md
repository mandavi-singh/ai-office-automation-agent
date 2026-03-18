# AI Office Automation Agent

A Windows desktop automation app for Word, Excel, PowerPoint, PDF, OCR, browser workflows, and AI-powered chat actions.

## Main App

The active application lives in `ai_agent_app/`.

Run it with:

```bash
cd ai_agent_app
python main.py
```

## Features

- OpenAI and Gemini agent support
- Word, Excel, PowerPoint, and PDF automation
- OCR tools
- Browser open/close/scrape tools
- Windows system app actions like Notepad and Calculator

## Setup

```bash
cd ai_agent_app
pip install -r requirements.txt
python main.py
```

## Notes

- Keep API keys in `ai_agent_app/.env`
- Local output and temp files are ignored by Git
