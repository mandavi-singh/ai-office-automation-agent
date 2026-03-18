"""
OpenAI agent with tool calling for office automation.
"""
import json
import re
from typing import Callable

from openai import OpenAI
from openai import APIError, AuthenticationError, BadRequestError, RateLimitError


TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "word_create_file",
            "description": "Create a new MS Word (.docx) file with a title and optional content.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "File path to save, e.g. C:/Reports/report.docx"},
                    "title": {"type": "string", "description": "Title text to add at top"},
                    "title_font_size": {"type": "integer", "description": "Font size for title (default 24)"},
                    "content": {"type": "string", "description": "Body content or paragraphs to add"},
                    "font_color": {"type": "string", "description": "Hex color for content text, e.g. FF0000"},
                },
                "required": ["filename", "title"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "word_format_text",
            "description": "Open an existing Word file and apply bold, italic, font color, or font size to matching text.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to the .docx file"},
                    "search_text": {"type": "string", "description": "Text to find and format"},
                    "bold": {"type": "boolean", "description": "Apply bold formatting"},
                    "italic": {"type": "boolean", "description": "Apply italic formatting"},
                    "font_color": {"type": "string", "description": "Hex color string e.g. FF0000"},
                    "font_size": {"type": "integer", "description": "Font size in points"},
                },
                "required": ["filename", "search_text"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "word_add_content",
            "description": "Add a new paragraph or heading to an existing Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to the .docx file"},
                    "text": {"type": "string", "description": "Text content to add"},
                    "style": {"type": "string", "description": "Style: Normal, Heading 1, Heading 2, Heading 3"},
                    "bold": {"type": "boolean", "description": "Apply bold"},
                    "italic": {"type": "boolean", "description": "Apply italic"},
                    "font_color": {"type": "string", "description": "Hex color e.g. 0070C0"},
                    "font_size": {"type": "integer", "description": "Font size in points"},
                },
                "required": ["filename", "text"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_create_table",
            "description": "Create a new Excel (.xlsx) file with a formatted table. Pass rows as a JSON string in rows_json.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to save the .xlsx file"},
                    "sheet_name": {"type": "string", "description": "Sheet name (default: Sheet1)"},
                    "headers": {"type": "array", "items": {"type": "string"}, "description": "Column header names"},
                    "rows_json": {"type": "string", "description": "Rows as JSON string"},
                },
                "required": ["filename"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_insert_value",
            "description": "Insert a specific value into a cell in an existing Excel file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to the .xlsx file"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "row": {"type": "integer", "description": "Row number (1-based)"},
                    "column": {"type": "integer", "description": "Column number (1-based)"},
                    "value": {"type": "string", "description": "Value to insert"},
                },
                "required": ["filename", "row", "column", "value"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_fill_demo",
            "description": "Create a demo Excel file with formatted sample data.",
            "parameters": {
                "type": "object",
                "properties": {"filename": {"type": "string", "description": "Path to save the .xlsx file"}},
                "required": ["filename"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_open_workbook",
            "description": "Open Microsoft Excel desktop, then open or create the requested workbook and optional sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Workbook path, or __ACTIVE__ to use the active workbook"},
                    "sheet_name": {"type": "string", "description": "Optional sheet name to activate or create"},
                    "visible": {"type": "boolean", "description": "Keep Excel visible to the user"},
                    "create_if_missing": {"type": "boolean", "description": "Create the workbook if it does not exist"},
                },
                "required": ["filename"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_describe_workbook",
            "description": "Inspect an Excel workbook and return sheet names plus used-range summary.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Workbook path, or __ACTIVE__ to use the active workbook"},
                    "sheet_name": {"type": "string", "description": "Optional sheet name to inspect"},
                },
                "required": ["filename"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_write_range",
            "description": "Write values into an open Excel workbook starting at a cell. Pass values as a JSON string in values_json.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Workbook path, or __ACTIVE__ to use the active workbook"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "start_cell": {"type": "string", "description": "Start cell like A1"},
                    "values_json": {"type": "string", "description": "2D JSON array string"},
                    "autofit": {"type": "boolean", "description": "Auto-fit columns after writing"},
                },
                "required": ["filename", "start_cell", "values_json"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_read_range",
            "description": "Read values from a range in an Excel workbook.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Workbook path, or __ACTIVE__ to use the active workbook"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "range_address": {"type": "string", "description": "Range like A1:C5"},
                },
                "required": ["filename", "range_address"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_format_range",
            "description": "Apply basic formatting such as bold, colors, autofit, and number formats to an Excel range.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Workbook path, or __ACTIVE__ to use the active workbook"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "range_address": {"type": "string", "description": "Range like A1:C5"},
                    "bold": {"type": "boolean", "description": "Apply bold font"},
                    "italic": {"type": "boolean", "description": "Apply italic font"},
                    "autofit": {"type": "boolean", "description": "Auto-fit columns"},
                    "font_size": {"type": "integer", "description": "Font size in points"},
                    "font_color": {"type": "string", "description": "Hex font color like FF0000"},
                    "fill_color": {"type": "string", "description": "Hex fill color like FFF2CC"},
                    "number_format": {"type": "string", "description": "Excel number format string"},
                    "border_style": {"type": "string", "description": "Excel border style value like 1 for continuous thin borders"},
                    "horizontal_alignment": {"type": "string", "description": "left, center, or right"},
                    "row_height": {"type": "number", "description": "Row height value such as 24"},
                },
                "required": ["filename", "range_address"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_save_workbook",
            "description": "Save an open Excel workbook.",
            "parameters": {
                "type": "object",
                "properties": {"filename": {"type": "string", "description": "Workbook path, or __ACTIVE__ to use the active workbook"}},
                "required": ["filename"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "excel_close_workbook",
            "description": "Close an Excel workbook after saving or discarding changes.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Workbook path, or __ACTIVE__ to use the active workbook"},
                    "save_changes": {"type": "boolean", "description": "Save changes before closing"},
                },
                "required": ["filename"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "ppt_create_presentation",
            "description": "Create a new PowerPoint (.pptx) file with a professional title slide.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to save .pptx file"},
                    "title": {"type": "string", "description": "Presentation title"},
                    "subtitle": {"type": "string", "description": "Subtitle or description"},
                    "company": {"type": "string", "description": "Company or author name"},
                    "date": {"type": "string", "description": "Date string"},
                    "title_font_size": {"type": "integer", "description": "Title font size in points"},
                    "subtitle_font_size": {"type": "integer", "description": "Subtitle font size in points"},
                    "title_font_color": {"type": "string", "description": "Title font color as hex, for example FF0000"},
                    "subtitle_font_color": {"type": "string", "description": "Subtitle font color as hex, for example FF0000"},
                    "background_color": {"type": "string", "description": "Slide background color as hex, for example FFFFFF"},
                    "fill_color": {"type": "string", "description": "Alias for background color as hex, for example FF0000"},
                    "title_font_name": {"type": "string", "description": "Title font family name, for example Aptos"},
                    "subtitle_font_name": {"type": "string", "description": "Subtitle font family name, for example Aptos"},
                },
                "required": ["filename", "title"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "ppt_add_slide",
            "description": "Add a new content slide with title and bullet points to an existing PowerPoint file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to the .pptx file"},
                    "title": {"type": "string", "description": "Slide title"},
                    "content": {"type": "array", "items": {"type": "string"}, "description": "List of bullet point strings"},
                    "layout": {"type": "string", "description": "title_content or title_only or blank"},
                    "title_font_size": {"type": "integer", "description": "Slide title font size in points"},
                    "content_font_size": {"type": "integer", "description": "Slide body font size in points"},
                    "title_font_color": {"type": "string", "description": "Slide title font color as hex"},
                    "content_font_color": {"type": "string", "description": "Slide body font color as hex"},
                    "background_color": {"type": "string", "description": "Slide background color as hex"},
                    "fill_color": {"type": "string", "description": "Alias for slide background color as hex"},
                    "title_font_name": {"type": "string", "description": "Slide title font family name, for example Aptos"},
                    "content_font_name": {"type": "string", "description": "Slide body font family name, for example Aptos"},
                },
                "required": ["filename", "title"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "ppt_add_table_slide",
            "description": "Add a native PowerPoint table slide with headers and data rows.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to the .pptx file"},
                    "title": {"type": "string", "description": "Slide title"},
                    "headers": {"type": "array", "items": {"type": "string"}, "description": "Table header names"},
                    "rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}, "description": "Table rows as a 2D array"},
                    "title_font_size": {"type": "integer", "description": "Slide title font size in points"},
                    "cell_font_size": {"type": "integer", "description": "Table cell font size in points"},
                    "title_font_color": {"type": "string", "description": "Slide title font color as hex"},
                    "cell_font_color": {"type": "string", "description": "Table cell font color as hex"},
                    "background_color": {"type": "string", "description": "Slide background color as hex"},
                    "fill_color": {"type": "string", "description": "Alias for slide background color as hex"},
                    "title_font_name": {"type": "string", "description": "Slide title font family name"},
                    "cell_font_name": {"type": "string", "description": "Table cell font family name"},
                },
                "required": ["filename", "title", "headers"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "ppt_edit_slide",
            "description": "Edit the title or content of an existing slide in a PowerPoint file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Path to .pptx file"},
                    "slide_index": {"type": "integer", "description": "0-based slide index"},
                    "new_title": {"type": "string", "description": "New title for the slide"},
                    "new_content": {"type": "array", "items": {"type": "string"}, "description": "New bullet points list"},
                    "title_font_size": {"type": "integer", "description": "Slide title font size in points"},
                    "content_font_size": {"type": "integer", "description": "Slide body font size in points"},
                    "title_font_color": {"type": "string", "description": "Slide title font color as hex"},
                    "content_font_color": {"type": "string", "description": "Slide body font color as hex"},
                    "title_font_name": {"type": "string", "description": "Slide title font family name, for example Algerian"},
                    "content_font_name": {"type": "string", "description": "Slide body font family name, for example Aptos"},
                },
                "required": ["filename", "slide_index"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "browser_open",
            "description": "Open Edge, Chrome, or Firefox with a URL or search query.",
            "parameters": {
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "Website URL to open"},
                    "browser": {"type": "string", "description": "Browser name: edge, chrome, or firefox"},
                    "search_query": {"type": "string", "description": "Optional search query to open in the browser"},
                    "interactive": {"type": "boolean", "description": "Open a Playwright-controlled browser window that can later scrape the current page"},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "browser_close",
            "description": "Close the last browser window opened by the app, or close all windows for a browser.",
            "parameters": {
                "type": "object",
                "properties": {
                    "browser": {"type": "string", "description": "Browser name: edge, chrome, or firefox"},
                    "close_all": {"type": "boolean", "description": "Set true to close all windows for that browser"},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "browser_scrape_current_page",
            "description": "Scrape the page currently open in the interactive automation browser.",
            "parameters": {
                "type": "object",
                "properties": {
                    "max_chars": {"type": "integer", "description": "Maximum text characters to return"},
                    "include_links": {"type": "boolean", "description": "Include extracted links in the response"},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "web_scrape_page",
            "description": "Fetch a webpage and return readable text plus links. Use render_js=true for JavaScript-rendered pages when Playwright is installed.",
            "parameters": {
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "Website URL to scrape"},
                    "max_chars": {"type": "integer", "description": "Maximum text characters to return"},
                    "include_links": {"type": "boolean", "description": "Include extracted links in the response"},
                    "render_js": {"type": "boolean", "description": "Render the page in a real browser before scraping"},
                    "browser": {"type": "string", "description": "Browser to use for rendered scraping: edge, chrome, or firefox"},
                    "wait_until": {"type": "string", "description": "Page load strategy such as load, domcontentloaded, or networkidle"},
                },
                "required": ["url"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "open_system_app",
            "description": "Open a supported Windows system app such as Notepad or Calculator.",
            "parameters": {
                "type": "object",
                "properties": {
                    "app_name": {"type": "string", "description": "System app name such as notepad or calculator"},
                },
                "required": ["app_name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "close_system_app",
            "description": "Close a supported Windows system app such as Notepad or Calculator.",
            "parameters": {
                "type": "object",
                "properties": {
                    "app_name": {"type": "string", "description": "System app name such as notepad or calculator"},
                },
                "required": ["app_name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "write_in_system_app",
            "description": "Write text into a supported system app such as Notepad.",
            "parameters": {
                "type": "object",
                "properties": {
                    "app_name": {"type": "string", "description": "System app name, currently notepad is supported"},
                    "text": {"type": "string", "description": "Text to type into the app"},
                },
                "required": ["app_name", "text"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "calculate_in_calculator",
            "description": "Open Calculator, enter a simple arithmetic expression, and calculate it.",
            "parameters": {
                "type": "object",
                "properties": {
                    "expression": {"type": "string", "description": "Arithmetic expression such as 4*8 or 25+17"},
                },
                "required": ["expression"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "pdf_extract_images",
            "description": "Extract and crop images from a PDF file and save them to an output folder.",
            "parameters": {
                "type": "object",
                "properties": {
                    "pdf_path": {"type": "string", "description": "Path to the PDF file"},
                    "output_folder": {"type": "string", "description": "Folder to save extracted images"},
                    "page_number": {"type": "integer", "description": "Page to extract from (1-based). 0 means all pages."},
                },
                "required": ["pdf_path", "output_folder"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "pdf_word_to_pdf",
            "description": "Convert a Word (.docx) file to PDF.",
            "parameters": {
                "type": "object",
                "properties": {
                    "docx_path": {"type": "string", "description": "Path to the .docx file"},
                    "pdf_path": {"type": "string", "description": "Output path for the PDF"},
                },
                "required": ["docx_path", "pdf_path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "pdf_extract_text",
            "description": "Extract text content from a PDF file preserving structure.",
            "parameters": {
                "type": "object",
                "properties": {
                    "pdf_path": {"type": "string", "description": "Path to the PDF file"},
                    "page_number": {"type": "integer", "description": "Page number (1-based). 0 = all pages."},
                },
                "required": ["pdf_path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "ocr_image_to_text",
            "description": "Extract text from an image using OCR and optionally save it as a text file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "image_path": {"type": "string", "description": "Path to the image file"},
                    "language": {"type": "string", "description": "OCR language such as eng, hin, or eng+hin"},
                    "save_to": {"type": "string", "description": "Optional output .txt file path"},
                },
                "required": ["image_path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "ocr_image_to_word",
            "description": "Extract text from an image using OCR and save it to a Word (.docx) file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "image_path": {"type": "string", "description": "Path to the image file"},
                    "output_docx": {"type": "string", "description": "Output path for the generated Word file"},
                    "language": {"type": "string", "description": "OCR language such as eng, hin, or eng+hin"},
                },
                "required": ["image_path", "output_docx"],
            },
        },
    },
]

SYSTEM_PROMPT = (
    "You are an intelligent Office Automation Agent. "
    "You help users automate MS Word, Excel, PowerPoint, PDF, and OCR tasks. "
    "Respond only in English. "
    "When the user asks to do work, call the appropriate function instead of only describing steps. "
    "Never say that you cannot physically open Excel or cannot open windows on the user's screen. "
    "For supported Windows desktop actions, use the system app tools instead of saying the user must do it manually. "
    "If the user asks to open Notepad, call open_system_app with notepad. "
    "If the user asks to write or type text in Notepad, call write_in_system_app with app_name notepad. "
    "If the user asks to open Calculator and calculate an expression, call calculate_in_calculator. "
    "If a user follows up with 'write ...' after opening Notepad, treat it as a request to type in Notepad. "
    "If a user follows up with 'calculate ...' after opening Calculator, treat it as a request to use Calculator. "
    "If the user asks to open Excel, create an Excel file, or work in Excel, you must use the available Excel tools. "
    "For excel_create_table, always pass rows in rows_json. "
    "For excel_write_range, always pass values in values_json. "
    "When the user asks to create an Excel file, make an Excel sheet, prepare an Excel report, or otherwise do Excel work, "
    "default to the live Excel desktop workflow instead of only file-based creation. "
    "That means you should usually call excel_open_workbook first, then use excel_write_range or other Excel tools, "
    "and finally save the workbook only when the user asked for saving or gave a file path. "
    "For Excel requests that mention opening Excel, editing interactively, or doing the whole job from a prompt, "
    "prefer the live Excel desktop tools: excel_open_workbook, excel_describe_workbook, excel_write_range, "
    "excel_format_range, excel_save_workbook, and excel_close_workbook. "
    "You may call multiple tools in sequence until the job is complete. "
    "If the user does not provide a filename for Excel, do not invent a save path. "
    "Instead, open Excel and use the active workbook by passing __ACTIVE__ as the filename argument. "
    "Always confirm the actual work completed after tools finish."
)


class OpenAIAgent:
    """OpenAI-based office automation agent with iterative tool calling."""

    def __init__(self, api_key: str, tool_executor: Callable, log_callback: Callable = None):
        self.client = OpenAI(api_key=api_key)
        self.tool_executor = tool_executor
        self.log = log_callback or print
        self.last_excel_filename = None
        self.last_system_app = None
        self.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
        self.model = "gpt-4.1-mini"

    def _remember_excel_target(self, fn_name: str, fn_args: dict):
        if not fn_name.startswith("excel_"):
            return
        filename = fn_args.get("filename")
        if filename and str(filename).strip().upper() != "__ACTIVE__":
            self.last_excel_filename = filename

    def _remember_system_app_target(self, fn_name: str, fn_args: dict):
        if fn_name not in {"open_system_app", "write_in_system_app"}:
            return
        app_name = str(fn_args.get("app_name", "")).strip().lower()
        if app_name in {"notepad", "calculator", "calc"}:
            self.last_system_app = "calculator" if app_name == "calc" else app_name

    def _is_notepad_request(self, user_message: str) -> bool:
        normalized = user_message.strip().lower()
        return "notepad" in normalized

    def _is_calculator_request(self, user_message: str) -> bool:
        normalized = user_message.strip().lower()
        return "calculator" in normalized or normalized.startswith("calc ")

    def _extract_notepad_text(self, user_message: str) -> str:
        message = user_message.strip()
        lowered = message.lower()

        for marker in ("write ", "type "):
            idx = lowered.find(marker)
            if idx != -1:
                return message[idx + len(marker):].strip().strip("\"'")
        return ""

    def _extract_calculator_expression(self, user_message: str) -> str:
        message = user_message.strip()
        lowered = message.lower()

        for marker in ("calculate ", "compute ", "evaluate "):
            idx = lowered.find(marker)
            if idx != -1:
                return message[idx + len(marker):].strip().strip("\"'")

        if re.fullmatch(r"[0-9\.\+\-\*/\(\) xX]+", message):
            return message
        return ""

    def _handle_direct_system_app_request(self, user_message: str) -> str | None:
        normalized = user_message.strip().lower()

        if self._is_notepad_request(user_message):
            text = self._extract_notepad_text(user_message)
            if text:
                open_result = self.tool_executor("open_system_app", {"app_name": "notepad"})
                self.log('[TOOL CALL] open_system_app({"app_name": "notepad"})')
                self.log(f"[TOOL RESULT] {open_result}")
                write_result = self.tool_executor("write_in_system_app", {"app_name": "notepad", "text": text})
                self.log(f'[TOOL CALL] write_in_system_app({json.dumps({"app_name": "notepad", "text": text}, indent=2)})')
                self.log(f"[TOOL RESULT] {write_result}")
                self.last_system_app = "notepad"
                return self._strip_non_english_lines(f"{open_result}\n{write_result}")

            if "open" in normalized:
                self.log('[TOOL CALL] open_system_app({"app_name": "notepad"})')
                result = self.tool_executor("open_system_app", {"app_name": "notepad"})
                self.log(f"[TOOL RESULT] {result}")
                self.last_system_app = "notepad"
                return self._strip_non_english_lines(result)

        if self._is_calculator_request(user_message):
            expr = self._extract_calculator_expression(user_message)
            if expr:
                self.log(f'[TOOL CALL] calculate_in_calculator({json.dumps({"expression": expr}, indent=2)})')
                result = self.tool_executor("calculate_in_calculator", {"expression": expr})
                self.log(f"[TOOL RESULT] {result}")
                self.last_system_app = "calculator"
                return self._strip_non_english_lines(result)

            if "open" in normalized:
                self.log('[TOOL CALL] open_system_app({"app_name": "calculator"})')
                result = self.tool_executor("open_system_app", {"app_name": "calculator"})
                self.log(f"[TOOL RESULT] {result}")
                self.last_system_app = "calculator"
                return self._strip_non_english_lines(result)

        if self.last_system_app == "notepad":
            text = self._extract_notepad_text(user_message)
            if text:
                self.log(f'[TOOL CALL] write_in_system_app({json.dumps({"app_name": "notepad", "text": text}, indent=2)})')
                result = self.tool_executor("write_in_system_app", {"app_name": "notepad", "text": text})
                self.log(f"[TOOL RESULT] {result}")
                return self._strip_non_english_lines(result)

        if self.last_system_app == "calculator":
            expr = self._extract_calculator_expression(user_message)
            if expr:
                self.log(f'[TOOL CALL] calculate_in_calculator({json.dumps({"expression": expr}, indent=2)})')
                result = self.tool_executor("calculate_in_calculator", {"expression": expr})
                self.log(f"[TOOL RESULT] {result}")
                return self._strip_non_english_lines(result)

        return None

    def _is_open_excel_request(self, user_message: str) -> bool:
        normalized = user_message.strip().lower()
        open_phrases = {
            "open it",
            "open this",
            "open the file",
            "open this file",
            "open excel",
            "open workbook",
            "open the workbook",
        }
        if normalized in open_phrases:
            return True
        return "open" in normalized and "excel" in normalized

    def _wants_excel_opened(self, user_message: str) -> bool:
        normalized = user_message.strip().lower()
        return "open it" in normalized or "open this" in normalized or ("open" in normalized and "excel" in normalized)

    def _strip_non_english_lines(self, text: str) -> str:
        cleaned_lines = []
        for line in text.splitlines():
            if re.search(r"[\u0900-\u097F]", line):
                continue
            cleaned_lines.append(line)
        cleaned = "\n".join(cleaned_lines).strip()
        return cleaned or "Done."

    def _extract_text(self, message) -> str:
        content = getattr(message, "content", "")
        if isinstance(content, str):
            return content.strip()

        text_parts = []
        for item in content or []:
            item_type = getattr(item, "type", "")
            if item_type == "text":
                text_value = getattr(item, "text", "")
                if isinstance(text_value, str):
                    text_parts.append(text_value)
                else:
                    nested = getattr(text_value, "value", "")
                    if nested:
                        text_parts.append(nested)
            elif isinstance(item, dict) and item.get("type") == "text":
                text_value = item.get("text", "")
                if isinstance(text_value, str):
                    text_parts.append(text_value)
                elif isinstance(text_value, dict):
                    text_parts.append(text_value.get("value", ""))

        extracted = "\n".join(part for part in text_parts if part).strip()
        return extracted or getattr(message, "refusal", "") or ""

    def _friendly_error(self, exc: Exception) -> str:
        if isinstance(exc, AuthenticationError):
            return "Your OpenAI API key is invalid. Please check the key in the .env file or in the sidebar."
        if isinstance(exc, RateLimitError):
            message = str(exc)
            if "insufficient_quota" in message:
                return "Your OpenAI API key is valid, but the account or project has no remaining quota. Please check billing and usage in the OpenAI dashboard."
            return "The OpenAI request hit a rate limit. Please wait a moment and try again."
        if isinstance(exc, BadRequestError):
            return f"OpenAI rejected the request: {exc}"
        if isinstance(exc, APIError):
            return f"OpenAI API error: {exc}"
        return str(exc)

    def send_message(self, user_message: str) -> str:
        self.log(f"[USER] {user_message}")
        direct_system_result = self._handle_direct_system_app_request(user_message)
        if direct_system_result:
            self.messages.append({"role": "user", "content": user_message})
            self.messages.append({"role": "assistant", "content": direct_system_result})
            self.log(f"[AI] {direct_system_result}")
            return direct_system_result

        if self._is_open_excel_request(user_message) and self.last_excel_filename:
            result = self.tool_executor(
                "excel_open_workbook",
                {"filename": self.last_excel_filename, "visible": True, "create_if_missing": False},
            )
            latest_text = self._strip_non_english_lines(f"I opened {self.last_excel_filename} in Excel.\n{result}")
            self.log(f"[AI] {latest_text}")
            return latest_text

        self.messages.append({"role": "user", "content": user_message})

        try:
            for _ in range(8):
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=self.messages,
                    tools=TOOLS,
                    tool_choice="auto",
                )
                assistant_message = response.choices[0].message

                if not assistant_message.tool_calls:
                    latest_text = self._extract_text(assistant_message)
                    if self._wants_excel_opened(user_message) and self.last_excel_filename:
                        open_result = self.tool_executor(
                            "excel_open_workbook",
                            {"filename": self.last_excel_filename, "visible": True, "create_if_missing": False},
                        )
                        latest_text = f"{latest_text}\nI also opened {self.last_excel_filename} in Excel.\n{open_result}"
                    latest_text = self._strip_non_english_lines(latest_text)
                    self.messages.append({"role": "assistant", "content": latest_text})
                    self.log(f"[AI] {latest_text}")
                    return latest_text

                self.messages.append(assistant_message.model_dump())

                for tool_call in assistant_message.tool_calls:
                    fn_name = tool_call.function.name
                    fn_args = json.loads(tool_call.function.arguments or "{}")

                    if "rows_json" in fn_args:
                        try:
                            fn_args["rows"] = json.loads(fn_args.pop("rows_json"))
                        except Exception:
                            fn_args["rows"] = []
                            fn_args.pop("rows_json", None)

                    self._remember_excel_target(fn_name, fn_args)
                    self._remember_system_app_target(fn_name, fn_args)
                    self.log(f"[TOOL CALL] {fn_name}({json.dumps(fn_args, indent=2, default=str)})")
                    result = self.tool_executor(fn_name, fn_args)
                    self.log(f"[TOOL RESULT] {result}")
                    self.messages.append(
                        {
                            "role": "tool",
                            "tool_call_id": tool_call.id,
                            "content": str(result),
                        }
                    )
        except Exception as exc:
            friendly = self._friendly_error(exc)
            self.log(f"[AI] {friendly}")
            return friendly

        fallback = "I could not complete the request after several tool calls."
        self.log(f"[AI] {fallback}")
        return fallback
