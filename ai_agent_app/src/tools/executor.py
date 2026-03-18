"""
Tool executor.
Routes AI tool calls to the correct automation tool.
"""
from src.tools.word_tools import word_create_file, word_add_content, word_format_text
from src.tools.excel_tools import (
    excel_close_workbook,
    excel_create_table,
    excel_describe_workbook,
    excel_fill_demo,
    excel_format_range,
    excel_insert_value,
    excel_open_workbook,
    excel_read_range,
    excel_save_workbook,
    excel_write_range,
)
from src.tools.ppt_tools import ppt_create_presentation, ppt_add_slide, ppt_add_table_slide, ppt_edit_slide
from src.tools.pdf_tools import pdf_extract_images, pdf_extract_text, pdf_word_to_pdf
from src.tools.ocr_tools import ocr_image_to_text, ocr_image_to_word
from src.tools.browser_tools import browser_open, browser_close, browser_scrape_current_page, web_scrape_page
from src.tools.system_tools import open_system_app, close_system_app, write_in_system_app, calculate_in_calculator


TOOL_MAP = {
    "word_create_file": word_create_file,
    "word_add_content": word_add_content,
    "word_format_text": word_format_text,
    "excel_create_table": excel_create_table,
    "excel_insert_value": excel_insert_value,
    "excel_fill_demo": excel_fill_demo,
    "excel_open_workbook": excel_open_workbook,
    "excel_describe_workbook": excel_describe_workbook,
    "excel_write_range": excel_write_range,
    "excel_read_range": excel_read_range,
    "excel_format_range": excel_format_range,
    "excel_save_workbook": excel_save_workbook,
    "excel_close_workbook": excel_close_workbook,
    "ppt_create_presentation": ppt_create_presentation,
    "ppt_add_slide": ppt_add_slide,
    "ppt_add_table_slide": ppt_add_table_slide,
    "ppt_edit_slide": ppt_edit_slide,
    "browser_open": browser_open,
    "browser_close": browser_close,
    "browser_scrape_current_page": browser_scrape_current_page,
    "web_scrape_page": web_scrape_page,
    "open_system_app": open_system_app,
    "close_system_app": close_system_app,
    "write_in_system_app": write_in_system_app,
    "calculate_in_calculator": calculate_in_calculator,
    "pdf_extract_images": pdf_extract_images,
    "pdf_extract_text": pdf_extract_text,
    "pdf_word_to_pdf": pdf_word_to_pdf,
    "ocr_image_to_text": ocr_image_to_text,
    "ocr_image_to_word": ocr_image_to_word,
}


def execute_tool(name: str, args: dict, raw: bool = False):
    fn = TOOL_MAP.get(name)
    if not fn:
        error = {"success": False, "message": f"Unknown tool: {name}"}
        return error if raw else f"ERROR: {error['message']}"

    try:
        result = fn(**args)
        if raw:
            if isinstance(result, dict):
                return result
            return {"success": True, "message": str(result), "text": ""}
        if isinstance(result, dict):
            prefix = "SUCCESS" if result.get("success") else "ERROR"
            return f"{prefix}: {result.get('message', 'Done')}"
        return str(result)
    except TypeError as exc:
        error = {"success": False, "message": f"Tool '{name}' wrong arguments: {exc}"}
        return error if raw else f"ERROR: {error['message']}"
    except Exception as exc:
        error = {"success": False, "message": f"Tool '{name}' failed: {exc}"}
        return error if raw else f"ERROR: {error['message']}"
