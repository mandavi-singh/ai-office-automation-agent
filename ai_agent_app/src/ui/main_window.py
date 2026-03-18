"""
Main Window — AI Agent for Office Automation
PySide6 Desktop Application
"""
import os
import json
import threading
import tempfile
from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QLineEdit, QPushButton, QLabel, QFrame,
    QScrollArea, QSplitter, QFileDialog, QTabWidget,
    QGroupBox, QFormLayout, QComboBox, QSpinBox,
    QCheckBox, QColorDialog, QStatusBar, QToolButton,
    QSizePolicy, QMessageBox, QApplication
)
from PySide6.QtCore import Qt, Signal, QObject, QThread, QSize
from PySide6.QtGui import QFont, QColor, QPalette, QIcon, QPixmap, QTextCursor, QShortcut, QKeySequence

from src.tools.executor import execute_tool


# ── STYLESHEET ────────────────────────────────────────────────────────────────
STYLESHEET = """
QMainWindow {
    background-color: #0F172A;
}

/* ── Sidebar ── */
#sidebar {
    background-color: #1A2B5E;
    min-width: 220px;
    max-width: 220px;
}
#sidebar QLabel#app_title {
    color: #FFFFFF;
    font-size: 16px;
    font-weight: bold;
    padding: 18px 16px 4px 16px;
}
#sidebar QLabel#app_sub {
    color: #7AADDC;
    font-size: 11px;
    padding: 0px 16px 16px 16px;
}
#sidebar QPushButton {
    background: transparent;
    color: #CADCFC;
    border: none;
    text-align: left;
    padding: 10px 18px;
    font-size: 13px;
    border-radius: 6px;
    margin: 2px 8px;
}
#sidebar QPushButton:hover {
    background-color: #1E3A7A;
    color: #FFFFFF;
}
#sidebar QPushButton[active="true"] {
    background-color: #00B4D8;
    color: #FFFFFF;
    font-weight: bold;
}
#sidebar_divider {
    background-color: #2A4A8A;
    max-height: 1px;
    margin: 6px 12px;
}
#api_status {
    color: #7AADDC;
    font-size: 10px;
    padding: 4px 16px;
}

/* ── Chat area ── */
#chat_area {
    background-color: #0F172A;
}
#chat_display {
    background-color: #111827;
    color: #E2E8F0;
    border: 1px solid #1E3A6A;
    border-radius: 8px;
    font-size: 13px;
    font-family: 'Segoe UI', Arial, sans-serif;
    padding: 12px;
    selection-background-color: #1E6FBA;
}
#chat_input {
    background-color: #1E293B;
    color: #F1F5F9;
    border: 1px solid #334155;
    border-radius: 8px;
    font-size: 14px;
    padding: 10px 14px;
    min-height: 42px;
}
#chat_input:focus {
    border: 1px solid #00B4D8;
}
#send_btn {
    background-color: #1E6FBA;
    color: white;
    border: none;
    border-radius: 8px;
    font-size: 14px;
    font-weight: bold;
    min-width: 90px;
    min-height: 42px;
    padding: 0 18px;
}
#send_btn:hover {
    background-color: #00B4D8;
}
#send_btn:disabled {
    background-color: #334155;
    color: #64748B;
}
#clear_btn {
    background-color: #1E293B;
    color: #94A3B8;
    border: 1px solid #334155;
    border-radius: 8px;
    font-size: 13px;
    min-height: 42px;
    padding: 0 14px;
}
#clear_btn:hover {
    background-color: #334155;
    color: #F1F5F9;
}

/* ── Tools panel ── */
#tools_panel {
    background-color: #0F172A;
}
#tool_group {
    color: #94A3B8;
    font-size: 12px;
    font-weight: bold;
    border: 1px solid #1E3A6A;
    border-radius: 8px;
    margin: 4px 0;
    padding-top: 8px;
}
#tool_group QLabel {
    color: #CBD5E1;
    font-size: 12px;
}
#tool_group QLineEdit, #tool_group QComboBox, #tool_group QSpinBox {
    background-color: #1E293B;
    color: #F1F5F9;
    border: 1px solid #334155;
    border-radius: 4px;
    padding: 5px 8px;
    font-size: 12px;
    min-height: 28px;
}
#tool_group QLineEdit:focus, #tool_group QComboBox:focus {
    border: 1px solid #00B4D8;
}
#tool_group QPushButton {
    background-color: #1E6FBA;
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 12px;
    font-weight: bold;
    padding: 7px 14px;
    margin-top: 4px;
}
#tool_group QPushButton:hover {
    background-color: #00B4D8;
}
#tool_group QPushButton#green_btn {
    background-color: #059669;
}
#tool_group QPushButton#green_btn:hover {
    background-color: #10B981;
}
#tool_group QPushButton#orange_btn {
    background-color: #D97706;
}
#tool_group QPushButton#orange_btn:hover {
    background-color: #F59E0B;
}
#tool_group QPushButton#red_btn {
    background-color: #B91C1C;
}
#tool_group QPushButton#red_btn:hover {
    background-color: #EF4444;
}
#tool_group QCheckBox {
    color: #CBD5E1;
    font-size: 12px;
}
#tool_group QCheckBox::indicator:checked {
    background-color: #00B4D8;
    border-radius: 3px;
}

/* ── API Key panel ── */
#api_panel {
    background-color: #111827;
    border: 1px solid #1E3A6A;
    border-radius: 8px;
    padding: 10px;
    margin: 8px;
}
#api_panel QLabel {
    color: #94A3B8;
    font-size: 12px;
}
#api_panel QLineEdit {
    background-color: #1E293B;
    color: #F1F5F9;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
}
#api_panel QPushButton {
    background-color: #059669;
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 12px;
    font-weight: bold;
    padding: 6px 16px;
}
#api_panel QPushButton:hover {
    background-color: #10B981;
}

/* ── Tabs ── */
QTabWidget::pane {
    background-color: #0F172A;
    border: 1px solid #1E3A6A;
    border-radius: 8px;
}
QTabBar::tab {
    background-color: #1E293B;
    color: #64748B;
    padding: 8px 20px;
    border: none;
    font-size: 12px;
    font-weight: bold;
}
QTabBar::tab:selected {
    background-color: #1E6FBA;
    color: #FFFFFF;
    border-radius: 6px 6px 0 0;
}
QTabBar::tab:hover {
    background-color: #334155;
    color: #E2E8F0;
}

/* ── Status bar ── */
QStatusBar {
    background-color: #1A2B5E;
    color: #7AADDC;
    font-size: 11px;
}
"""


# ── WORKER THREAD ─────────────────────────────────────────────────────────────
class AgentWorker(QObject):
    finished = Signal(str)
    error = Signal(str)
    log = Signal(str)

    def __init__(self, agent, message):
        super().__init__()
        self.agent = agent
        self.message = message

    def run(self):
        try:
            result = self.agent.send_message(self.message)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


# ── CHAT BUBBLE HELPERS ───────────────────────────────────────────────────────
def user_bubble(text: str) -> str:
    return (
        f'<div style="margin: 8px 0; text-align: right;">'
        f'<span style="display:inline-block; background:#1E6FBA; color:#FFFFFF; '
        f'padding:10px 14px; border-radius:14px 14px 2px 14px; max-width:75%; '
        f'font-size:13px; line-height:1.5;">{text}</span></div>'
    )


def ai_bubble(text: str) -> str:
    # Convert newlines and basic markdown
    text = text.replace("\n", "<br>")
    return (
        f'<div style="margin: 8px 0;">'
        f'<span style="display:inline-block; background:#1E293B; color:#E2E8F0; '
        f'border-left:3px solid #00B4D8; '
        f'padding:10px 14px; border-radius:2px 14px 14px 14px; max-width:85%; '
        f'font-size:13px; line-height:1.6;">{text}</span></div>'
    )


def tool_bubble(text: str) -> str:
    text = text.replace("\n", "<br>")
    color = "#10B981" if "✅" in text else ("#F59E0B" if "⚠️" in text else "#EF4444")
    return (
        f'<div style="margin: 4px 0 4px 16px;">'
        f'<span style="display:inline-block; background:#0F2A1A; color:{color}; '
        f'border:1px solid {color}44; '
        f'padding:6px 12px; border-radius:6px; '
        f'font-size:11px; font-family:Consolas,monospace;">{text}</span></div>'
    )


def thinking_bubble() -> str:
    return (
        '<div style="margin: 8px 0;">'
        '<span style="display:inline-block; background:#1E293B; color:#64748B; '
        'padding:10px 14px; border-radius:2px 14px 14px 14px; '
        'font-size:13px; font-style:italic;">⏳ Agent is thinking...</span></div>'
    )


# ── MAIN WINDOW ───────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self, api_key: str = ""):
        super().__init__()
        self.setWindowTitle("AI Agent — Office Automation")
        self.resize(1280, 800)
        self.setStyleSheet(STYLESHEET)
        self.agent = None
        self._env_api_key = api_key
        self._thinking_marker = None

        self._build_ui()
        self._status("Ready — Enter your OpenAI API key to start")

        # Auto-fill from .env if key found
        if self._env_api_key and self._env_api_key != "your_api_key_here":
            self.api_input.setText(self._env_api_key)
            self._connect_agent()

    def _save_openai_key_to_env(self, api_key: str):
        env_path = Path(__file__).resolve().parents[2] / ".env"
        existing_lines = []
        if env_path.exists():
            existing_lines = env_path.read_text(encoding="utf-8").splitlines()

        updated = []
        replaced = False
        for line in existing_lines:
            if line.strip().startswith("OPENAI_API_KEY="):
                updated.append(f"OPENAI_API_KEY={api_key}")
                replaced = True
            else:
                updated.append(line)

        if not replaced:
            if updated and updated[-1].strip():
                updated.append("")
            updated.append(f"OPENAI_API_KEY={api_key}")

        env_path.write_text("\n".join(updated).rstrip() + "\n", encoding="utf-8")

    # ── BUILD UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Sidebar
        sidebar = self._build_sidebar()
        layout.addWidget(sidebar)

        # Main content splitter
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(1)
        splitter.setStyleSheet("QSplitter::handle { background: #1E3A6A; }")

        splitter.addWidget(self._build_chat_area())
        splitter.addWidget(self._build_tools_panel())
        splitter.setSizes([820, 460])

        layout.addWidget(splitter)

        # Status bar
        self.statusBar().setObjectName("status_bar")
        self.statusBar().showMessage("Ready")

    def _build_sidebar(self):
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Title
        title = QLabel("🤖 AI Agent")
        title.setObjectName("app_title")
        sub = QLabel("Office Automation")
        sub.setObjectName("app_sub")
        layout.addWidget(title)
        layout.addWidget(sub)

        # Divider
        div = QFrame()
        div.setObjectName("sidebar_divider")
        div.setMaximumHeight(1)
        layout.addWidget(div)

        # Nav buttons
        self.nav_btns = {}
        nav_items = [
            ("💬 Chat",       "chat"),
            ("📝 Word",       "word"),
            ("📊 Excel",      "excel"),
            ("📊 PowerPoint", "ppt"),
            ("📄 PDF",        "pdf"),
        ]
        for label, key in nav_items:
            btn = QPushButton(label)
            btn.setProperty("active", key == "chat")
            btn.clicked.connect(lambda checked, k=key: self._nav_click(k))
            self.nav_btns[key] = btn
            layout.addWidget(btn)

        layout.addStretch()

        # API key section
        api_frame = QFrame()
        api_frame.setObjectName("api_panel")
        api_layout = QVBoxLayout(api_frame)
        api_layout.setSpacing(6)

        api_lbl = QLabel("🔑 OpenAI API Key")
        api_lbl.setStyleSheet("color:#94A3B8; font-size:11px; font-weight:bold;")
        self.api_input = QLineEdit()
        self.api_input.setPlaceholderText("Auto-loaded from .env (or type here)")
        self.api_input.setEchoMode(QLineEdit.Password)

        connect_btn = QPushButton("Connect")
        connect_btn.clicked.connect(self._connect_agent)

        self.api_status_lbl = QLabel("Checking .env...")
        self.api_status_lbl.setObjectName("api_status")

        api_layout.addWidget(api_lbl)
        api_layout.addWidget(self.api_input)
        api_layout.addWidget(connect_btn)
        api_layout.addWidget(self.api_status_lbl)
        layout.addWidget(api_frame)

        layout.addSpacing(8)
        return sidebar

    def _build_chat_area(self):
        frame = QFrame()
        frame.setObjectName("chat_area")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 16, 8, 16)
        layout.setSpacing(10)

        # Header
        header_lbl = QLabel("💬 Chat with AI Agent")
        header_lbl.setStyleSheet("color:#E2E8F0; font-size:15px; font-weight:bold;")
        layout.addWidget(header_lbl)

        # Chat display
        self.chat_display = QTextEdit()
        self.chat_display.setObjectName("chat_display")
        self.chat_display.setReadOnly(True)
        self.chat_display.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.chat_display)

        # Suggestions row
        suggestions_layout = QHBoxLayout()
        suggestions_layout.setSpacing(8)
        suggestions = [
            "Create a Word file",
            "Open Excel and create a sales sheet",
            "Create a PowerPoint",
            "Extract PDF images",
        ]
        for s in suggestions:
            btn = QPushButton(s)
            btn.setStyleSheet(
                "background:#1E293B; color:#94A3B8; border:1px solid #334155;"
                "border-radius:12px; padding:4px 12px; font-size:11px;"
            )
            btn.clicked.connect(lambda checked, text=s: self._set_input(text))
            suggestions_layout.addWidget(btn)
        suggestions_layout.addStretch()
        layout.addLayout(suggestions_layout)

        # Input row
        input_row = QHBoxLayout()
        input_row.setSpacing(8)

        self.chat_input = QLineEdit()
        self.chat_input.setObjectName("chat_input")
        self.chat_input.setPlaceholderText("Type a command... e.g. 'Open Excel and create a March sales sheet'")
        self.chat_input.returnPressed.connect(self._send_message)
        input_row.addWidget(self.chat_input)

        clear_btn = QPushButton("Clear")
        clear_btn.setObjectName("clear_btn")
        clear_btn.clicked.connect(self._clear_chat)
        input_row.addWidget(clear_btn)

        self.send_btn = QPushButton("Send ▶")
        self.send_btn.setObjectName("send_btn")
        self.send_btn.clicked.connect(self._send_message)
        input_row.addWidget(self.send_btn)

        layout.addLayout(input_row)

        # Welcome message
        self.chat_display.append(ai_bubble(
            "👋 Hello! I'm your <b>AI Office Automation Agent</b>.<br><br>"
            "I can help you automate with natural-language prompts:<br>"
            "• 📝 <b>MS Word</b> — create files, bold/italic text, set font colors<br>"
            "• 📊 <b>Excel</b> — open Excel, create/edit sheets, write data, format ranges<br>"
            "• 📊 <b>PowerPoint</b> — create slides, add content, edit existing files<br>"
            "• 📄 <b>PDF</b> — extract images, extract text, convert Word → PDF<br><br>"
            "Connect your <b>OpenAI API key</b> in the sidebar, then tell me what to do. For example: "
            "<i>'Open Excel and create an attendance sheet'</i> 🚀"
        ))
        return frame

    def _build_tools_panel(self):
        frame = QFrame()
        frame.setObjectName("tools_panel")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(8, 16, 16, 16)
        layout.setSpacing(0)

        header_lbl = QLabel("🛠️ Quick Actions")
        header_lbl.setStyleSheet("color:#E2E8F0; font-size:15px; font-weight:bold; margin-bottom:8px;")
        layout.addWidget(header_lbl)

        tabs = QTabWidget()
        tabs.addTab(self._wrap_tab_content(self._build_word_tab()),  "📝 Word")
        tabs.addTab(self._wrap_tab_content(self._build_excel_tab()), "📊 Excel")
        tabs.addTab(self._wrap_tab_content(self._build_ppt_tab()),   "📊 PPT")
        tabs.addTab(self._wrap_tab_content(self._build_pdf_tab()),   "📄 PDF")
        layout.addWidget(tabs)

        browser_help = QGroupBox("Browser Help")
        browser_help.setObjectName("tool_group")
        browser_help_layout = QVBoxLayout(browser_help)
        browser_help_layout.setContentsMargins(12, 12, 12, 12)
        browser_help_layout.setSpacing(6)

        browser_help_text = QLabel(
            "Use these prompts for browser tools:<br>"
            "• <b>Open Edge and go to google.com</b><br>"
            "• <b>Scrape https://example.com and return up to 500 characters</b><br>"
            "• <b>Scrape https://www.python.org with JavaScript rendering using Edge</b><br>"
            "• <b>Open Edge interactively and go to https://openai.com</b><br>"
            "• <b>Scrape the current browser page</b><br>"
            "• <b>Close the browser</b><br><br>"
            "<span style='color:#94A3B8;'>Interactive scraping works best after you manually complete login or verification steps.</span>"
        )
        browser_help_text.setWordWrap(True)
        browser_help_text.setStyleSheet("color:#E2E8F0; font-size:12px;")
        browser_help_layout.addWidget(browser_help_text)

        layout.addWidget(browser_help)
        return frame

    # ── WORD TAB ──────────────────────────────────────────────────────────────
    def _build_word_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # Create file
        grp = QGroupBox("Create Word File")
        grp.setObjectName("tool_group")
        fl = QFormLayout(grp)
        self.w_filename = QLineEdit("C:/output/report.docx")
        self.w_title = QLineEdit("My Report")
        self.w_title_size = QSpinBox()
        self.w_title_size.setRange(10, 72)
        self.w_title_size.setValue(24)
        self.w_content = QTextEdit()
        self.w_content.setMaximumHeight(60)
        self.w_content.setPlaceholderText("Body text...")
        self.w_content.setStyleSheet("background:#1E293B; color:#F1F5F9; border:1px solid #334155; border-radius:4px; font-size:12px;")
        self.w_font_color = QLineEdit("000000")
        self.w_font_color.setPlaceholderText("Hex e.g. FF0000")
        fl.addRow("Filename:", self.w_filename)
        fl.addRow("Title:", self.w_title)
        fl.addRow("Title Size:", self.w_title_size)
        fl.addRow("Content:", self.w_content)
        fl.addRow("Font Color:", self.w_font_color)
        btn_create = QPushButton("📝 Create Word File")
        btn_create.clicked.connect(self._direct_word_create)
        fl.addRow(btn_create)
        layout.addWidget(grp)

        # Format text
        grp2 = QGroupBox("Format Text in Existing File")
        grp2.setObjectName("tool_group")
        fl2 = QFormLayout(grp2)
        self.w_fmt_file = QLineEdit()
        self.w_fmt_file.setPlaceholderText("Path to .docx...")
        self.w_fmt_browse = QPushButton("Browse")
        self.w_fmt_browse.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        self.w_fmt_browse.clicked.connect(lambda: self._browse_file(self.w_fmt_file, "Word Files (*.docx)"))
        self.w_fmt_search = QLineEdit()
        self.w_fmt_search.setPlaceholderText("Text to find and format")
        self.w_fmt_bold = QCheckBox("Bold")
        self.w_fmt_italic = QCheckBox("Italic")
        self.w_fmt_color = QLineEdit()
        self.w_fmt_color.setPlaceholderText("Hex color e.g. FF0000")
        fl2.addRow("File:", self._row(self.w_fmt_file, self.w_fmt_browse))
        fl2.addRow("Search:", self.w_fmt_search)
        fl2.addRow("Style:", self._row(self.w_fmt_bold, self.w_fmt_italic))
        fl2.addRow("Color:", self.w_fmt_color)
        btn_fmt = QPushButton("✏️ Apply Formatting")
        btn_fmt.clicked.connect(self._direct_word_format)
        fl2.addRow(btn_fmt)
        layout.addWidget(grp2)

        layout.addStretch()
        return w

    # ── EXCEL TAB ─────────────────────────────────────────────────────────────
    def _build_excel_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # Demo
        grp = QGroupBox("Demo — Values 1 to 6")
        grp.setObjectName("tool_group")
        fl = QFormLayout(grp)
        self.xl_demo_file = QLineEdit("C:/output/demo_excel.xlsx")
        fl.addRow("Filename:", self.xl_demo_file)
        btn = QPushButton("📊 Create Demo Excel")
        btn.setObjectName("green_btn")
        btn.clicked.connect(self._direct_excel_demo)
        fl.addRow(btn)
        layout.addWidget(grp)

        # Create table
        grp2 = QGroupBox("Create Custom Table")
        grp2.setObjectName("tool_group")
        fl2 = QFormLayout(grp2)
        self.xl_filename = QLineEdit("C:/output/table.xlsx")
        self.xl_headers = QLineEdit("Name, Value, Status")
        self.xl_rows = QTextEdit()
        self.xl_rows.setMaximumHeight(70)
        self.xl_rows.setPlaceholderText("Row 1: Alice, 100, Active\nRow 2: Bob, 200, Pending")
        self.xl_rows.setStyleSheet("background:#1E293B; color:#F1F5F9; border:1px solid #334155; border-radius:4px; font-size:12px;")
        fl2.addRow("Filename:", self.xl_filename)
        fl2.addRow("Headers:", self.xl_headers)
        fl2.addRow("Rows:", self.xl_rows)
        btn2 = QPushButton("📊 Create Table")
        btn2.clicked.connect(self._direct_excel_table)
        fl2.addRow(btn2)
        layout.addWidget(grp2)

        # Insert value
        grp3 = QGroupBox("Insert Value into Cell")
        grp3.setObjectName("tool_group")
        fl3 = QFormLayout(grp3)
        self.xl_ins_file = QLineEdit()
        self.xl_ins_file.setPlaceholderText("Path to .xlsx...")
        browse3 = QPushButton("Browse")
        browse3.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        browse3.clicked.connect(lambda: self._browse_file(self.xl_ins_file, "Excel Files (*.xlsx)"))
        self.xl_ins_row = QSpinBox(); self.xl_ins_row.setRange(1, 1000); self.xl_ins_row.setValue(2)
        self.xl_ins_col = QSpinBox(); self.xl_ins_col.setRange(1, 100); self.xl_ins_col.setValue(1)
        self.xl_ins_val = QLineEdit("100")
        fl3.addRow("File:", self._row(self.xl_ins_file, browse3))
        fl3.addRow("Row:", self.xl_ins_row)
        fl3.addRow("Column:", self.xl_ins_col)
        fl3.addRow("Value:", self.xl_ins_val)
        btn3 = QPushButton("📥 Insert Value")
        btn3.clicked.connect(self._direct_excel_insert)
        fl3.addRow(btn3)
        layout.addWidget(grp3)

        grp4 = QGroupBox("OCR from Image")
        grp4.setObjectName("tool_group")
        fl4 = QFormLayout(grp4)
        self.ocr_image_path = QLineEdit()
        self.ocr_image_path.setPlaceholderText("Browse an image or paste one from clipboard...")
        self.ocr_image_path.textChanged.connect(self._update_ocr_preview)
        browse4 = QPushButton("Browse")
        browse4.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        browse4.clicked.connect(lambda: self._browse_file(
            self.ocr_image_path,
            "Image Files (*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.gif *.webp)"
        ))
        paste4 = QPushButton("Paste Image")
        paste4.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        paste4.clicked.connect(self._paste_ocr_image)
        fl4.addRow("Image:", self._row(self.ocr_image_path, browse4, paste4))

        self.ocr_lang = QComboBox()
        self.ocr_lang.addItems(["eng", "hin", "eng+hin"])
        fl4.addRow("Language:", self.ocr_lang)

        self.ocr_save_txt = QLineEdit()
        self.ocr_save_txt.setPlaceholderText("Optional output .txt path")
        browse_txt = QPushButton("Save As")
        browse_txt.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        browse_txt.clicked.connect(lambda: self._browse_save_file(self.ocr_save_txt, "Text Files (*.txt)"))
        fl4.addRow("TXT Output:", self._row(self.ocr_save_txt, browse_txt))

        self.ocr_save_docx = QLineEdit()
        self.ocr_save_docx.setPlaceholderText("Optional output .docx path")
        browse_docx = QPushButton("Save As")
        browse_docx.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        browse_docx.clicked.connect(lambda: self._browse_save_file(self.ocr_save_docx, "Word Files (*.docx)"))
        fl4.addRow("DOCX Output:", self._row(self.ocr_save_docx, browse_docx))

        self.ocr_preview = QLabel("Image preview will appear here")
        self.ocr_preview.setAlignment(Qt.AlignCenter)
        self.ocr_preview.setMinimumHeight(160)
        self.ocr_preview.setStyleSheet("background:#111827; color:#64748B; border:1px dashed #334155; border-radius:6px; padding:8px;")
        fl4.addRow("Preview:", self.ocr_preview)

        ocr_actions = QHBoxLayout()
        btn4 = QPushButton("Extract Text")
        btn4.clicked.connect(self._direct_ocr_extract_text)
        btn5 = QPushButton("Save to Word")
        btn5.setObjectName("green_btn")
        btn5.clicked.connect(self._direct_ocr_to_word)
        btn6 = QPushButton("Save Result")
        btn6.setObjectName("orange_btn")
        btn6.clicked.connect(self._save_ocr_result)
        ocr_actions.addWidget(btn4)
        ocr_actions.addWidget(btn5)
        ocr_actions.addWidget(btn6)
        fl4.addRow(ocr_actions)

        self.ocr_result = QTextEdit()
        self.ocr_result.setPlaceholderText("Extracted OCR text will appear here...")
        self.ocr_result.setMinimumHeight(150)
        self.ocr_result.setStyleSheet("background:#1E293B; color:#F1F5F9; border:1px solid #334155; border-radius:4px; font-size:12px; padding:6px;")
        fl4.addRow("Result:", self.ocr_result)

        ocr_hint = QLabel("Tip: copy any image, open this tab, and press Ctrl+V or click Paste Image.")
        ocr_hint.setStyleSheet("color:#94A3B8; font-size:11px;")
        fl4.addRow(ocr_hint)
        layout.addWidget(grp4)

        self.ocr_paste_shortcut = QShortcut(QKeySequence.Paste, grp4)
        self.ocr_paste_shortcut.activated.connect(self._paste_ocr_image)

        layout.addStretch()
        return w

    # ── PPT TAB ───────────────────────────────────────────────────────────────
    def _build_ppt_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # Create
        grp = QGroupBox("Create Presentation")
        grp.setObjectName("tool_group")
        fl = QFormLayout(grp)
        self.ppt_filename = QLineEdit("C:/output/presentation.pptx")
        self.ppt_title = QLineEdit("My Presentation")
        self.ppt_subtitle = QLineEdit("Subtitle here")
        self.ppt_company = QLineEdit("My Company")
        self.ppt_date = QLineEdit()
        self.ppt_date.setPlaceholderText("March 2026 (optional)")
        fl.addRow("Filename:", self.ppt_filename)
        fl.addRow("Title:", self.ppt_title)
        fl.addRow("Subtitle:", self.ppt_subtitle)
        fl.addRow("Company:", self.ppt_company)
        fl.addRow("Date:", self.ppt_date)
        btn = QPushButton("📊 Create Presentation")
        btn.setObjectName("orange_btn")
        btn.clicked.connect(self._direct_ppt_create)
        fl.addRow(btn)
        layout.addWidget(grp)

        # Add slide
        grp2 = QGroupBox("Add Slide to Existing File")
        grp2.setObjectName("tool_group")
        fl2 = QFormLayout(grp2)
        self.ppt_add_file = QLineEdit()
        self.ppt_add_file.setPlaceholderText("Path to .pptx...")
        browse2 = QPushButton("Browse")
        browse2.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        browse2.clicked.connect(lambda: self._browse_file(self.ppt_add_file, "PowerPoint Files (*.pptx)"))
        self.ppt_add_title = QLineEdit("New Slide Title")
        self.ppt_add_content = QTextEdit()
        self.ppt_add_content.setMaximumHeight(70)
        self.ppt_add_content.setPlaceholderText("Point 1\nPoint 2\nPoint 3")
        self.ppt_add_content.setStyleSheet("background:#1E293B; color:#F1F5F9; border:1px solid #334155; border-radius:4px; font-size:12px;")
        fl2.addRow("File:", self._row(self.ppt_add_file, browse2))
        fl2.addRow("Title:", self.ppt_add_title)
        fl2.addRow("Bullets:", self.ppt_add_content)
        btn2 = QPushButton("➕ Add Slide")
        btn2.clicked.connect(self._direct_ppt_add_slide)
        fl2.addRow(btn2)
        layout.addWidget(grp2)

        layout.addStretch()
        return w

    # ── PDF TAB ───────────────────────────────────────────────────────────────
    def _build_pdf_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # Extract text
        grp2 = QGroupBox("Extract Text from PDF")
        grp2.setObjectName("tool_group")
        fl2 = QFormLayout(grp2)
        self.pdf_txt_src = QLineEdit()
        self.pdf_txt_src.setPlaceholderText("Path to PDF...")
        browse2 = QPushButton("Browse")
        browse2.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        browse2.clicked.connect(lambda: self._browse_file(self.pdf_txt_src, "PDF Files (*.pdf)"))
        self.pdf_txt_page = QSpinBox()
        self.pdf_txt_page.setRange(0, 9999)
        self.pdf_txt_page.setValue(0)
        self.pdf_txt_page.setSpecialValueText("All Pages")
        fl2.addRow("PDF File:", self._row(self.pdf_txt_src, browse2))
        fl2.addRow("Page (0=All):", self.pdf_txt_page)
        btn2 = QPushButton("📖 Extract Text")
        btn2.clicked.connect(self._direct_pdf_extract_text)
        fl2.addRow(btn2)
        layout.addWidget(grp2)

        # Word to PDF
        grp3 = QGroupBox("Convert Word → PDF")
        grp3.setObjectName("tool_group")
        fl3 = QFormLayout(grp3)
        self.pdf_docx = QLineEdit()
        self.pdf_docx.setPlaceholderText("Path to .docx...")
        browse3 = QPushButton("Browse")
        browse3.setStyleSheet("background:#1E293B; color:#94A3B8; border:1px solid #334155; border-radius:4px; padding:4px 8px; font-size:11px;")
        browse3.clicked.connect(lambda: self._browse_file(self.pdf_docx, "Word Files (*.docx)"))
        self.pdf_out_pdf = QLineEdit(os.path.join(os.getcwd(), "output", "output.pdf"))
        fl3.addRow("Word File:", self._row(self.pdf_docx, browse3))
        fl3.addRow("PDF Output:", self.pdf_out_pdf)
        btn3 = QPushButton("🔄 Convert to PDF")
        btn3.clicked.connect(self._direct_word_to_pdf)
        fl3.addRow(btn3)
        layout.addWidget(grp3)

        layout.addStretch()
        return w

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _row(self, *widgets):
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(6)
        for index, widget in enumerate(widgets):
            if index == 0:
                h.addWidget(widget, 1)
            else:
                h.addWidget(widget, 0)
        return w

    def _wrap_tab_content(self, widget: QWidget):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(
            "QScrollArea { background: #0F172A; border: none; }"
            "QScrollArea > QWidget > QWidget { background: #0F172A; }"
        )
        widget.setStyleSheet("background: #0F172A;")
        scroll.setWidget(widget)
        return scroll

    def _browse_file(self, line_edit: QLineEdit, filter_str: str):
        path, _ = QFileDialog.getOpenFileName(self, "Select File", "", filter_str)
        if path:
            line_edit.setText(path)

    def _browse_save_file(self, line_edit: QLineEdit, filter_str: str):
        path, _ = QFileDialog.getSaveFileName(self, "Save File", line_edit.text().strip(), filter_str)
        if path:
            line_edit.setText(path)

    def _set_ocr_image_path(self, path: str):
        self.ocr_image_path.setText(path)
        self._update_ocr_preview(path)

    def _update_ocr_preview(self, path: str):
        if not path or not os.path.exists(path):
            self.ocr_preview.setText("Image preview will appear here")
            self.ocr_preview.setPixmap(QPixmap())
            return

        pixmap = QPixmap(path)
        if pixmap.isNull():
            self.ocr_preview.setText("Preview unavailable for this image")
            self.ocr_preview.setPixmap(QPixmap())
            return

        self.ocr_preview.setPixmap(
            pixmap.scaled(280, 180, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        )
        self.ocr_preview.setText("")

    def _paste_ocr_image(self):
        clipboard = QApplication.clipboard()
        pixmap = clipboard.pixmap()
        if pixmap.isNull():
            self.ocr_result.setPlainText("No image found in clipboard. Copy an image first, then try again.")
            return

        temp_dir = os.path.join(tempfile.gettempdir(), "ai_agent_ocr")
        os.makedirs(temp_dir, exist_ok=True)
        output_path = os.path.join(temp_dir, "clipboard_image.png")
        pixmap.save(output_path, "PNG")
        self._set_ocr_image_path(output_path)
        self.ocr_result.setPlainText(f"Clipboard image loaded.\nTemporary file: {output_path}")

    def _set_input(self, text: str):
        self.chat_input.setText(text)
        self.chat_input.setFocus()

    def _status(self, msg: str):
        self.statusBar().showMessage(msg)

    def _nav_click(self, key: str):
        for k, btn in self.nav_btns.items():
            btn.setProperty("active", k == key)
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    # ── AGENT CONNECTION ──────────────────────────────────────────────────────
    def _connect_agent(self):
        api_key = self.api_input.text().strip()
        if not api_key:
            self._append_chat(tool_bubble("❌ Please enter an OpenAI API key."))
            return
        try:
            self._save_openai_key_to_env(api_key)
            from src.agents.openai_agent import OpenAIAgent

            def log_fn(msg):
                if msg.startswith("[TOOL"):
                    self._append_chat(tool_bubble(msg))

            self.agent = OpenAIAgent(
                api_key=api_key,
                tool_executor=execute_tool,
                log_callback=log_fn,
            )
            self.api_status_lbl.setText("✅ Connected")
            self.api_status_lbl.setStyleSheet("color:#10B981; font-size:10px; padding:4px 16px;")
            self._append_chat(ai_bubble("✅ <b>Connected to OpenAI!</b> Your API key is saved in <b>.env</b> and will auto-load next time. You can now give direct prompts, for example: <i>Open Excel and create a budget tracker</i>."))
            self._status("Connected to OpenAI")
        except Exception as e:
            self.api_status_lbl.setText("❌ Failed")
            self._append_chat(tool_bubble(f"❌ Connection failed: {e}"))

    # ── CHAT ──────────────────────────────────────────────────────────────────
    def _append_chat(self, html: str):
        self.chat_display.append(html)
        self.chat_display.moveCursor(QTextCursor.End)

    def _clear_chat(self):
        self.chat_display.clear()

    def _send_message(self):
        text = self.chat_input.text().strip()
        if not text:
            return
        if not self.agent:
            self._append_chat(tool_bubble("⚠️ Please connect your OpenAI API key first."))
            return

        self.chat_input.clear()
        self._append_chat(user_bubble(text))
        self._append_chat(thinking_bubble())
        self.send_btn.setEnabled(False)
        self._status("Agent is processing...")

        # Run in thread
        self._worker = AgentWorker(self.agent, text)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_agent_done)
        self._worker.error.connect(self._on_agent_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.start()

    def _remove_thinking(self):
        """Remove the last thinking bubble."""
        cursor = self.chat_display.textCursor()
        cursor.movePosition(QTextCursor.End)
        # Simple approach: just append new content — thinking disappears visually
        # For a proper removal we'd track the position. Leaving as-is for clarity.

    def _on_agent_done(self, result: str):
        self._append_chat(ai_bubble(result))
        self.send_btn.setEnabled(True)
        self._status("Ready")

    def _on_agent_error(self, error: str):
        self._append_chat(tool_bubble(f"❌ Agent error: {error}"))
        self.send_btn.setEnabled(True)
        self._status("Error occurred")

    # ── DIRECT TOOL BUTTONS ───────────────────────────────────────────────────
    def _run_tool(self, fn_name: str, args: dict):
        result = execute_tool(fn_name, args)
        self._append_chat(tool_bubble(result))
        self._status(result[:80])

    def _direct_word_create(self):
        self._run_tool("word_create_file", {
            "filename": self.w_filename.text(),
            "title": self.w_title.text(),
            "title_font_size": self.w_title_size.value(),
            "content": self.w_content.toPlainText(),
            "font_color": self.w_font_color.text(),
        })

    def _direct_word_format(self):
        self._run_tool("word_format_text", {
            "filename": self.w_fmt_file.text(),
            "search_text": self.w_fmt_search.text(),
            "bold": self.w_fmt_bold.isChecked(),
            "italic": self.w_fmt_italic.isChecked(),
            "font_color": self.w_fmt_color.text(),
        })

    def _direct_excel_demo(self):
        self._run_tool("excel_fill_demo", {"filename": self.xl_demo_file.text()})

    def _direct_excel_table(self):
        headers = [h.strip() for h in self.xl_headers.text().split(",") if h.strip()]
        rows = []
        for line in self.xl_rows.toPlainText().strip().split("\n"):
            if line.strip():
                cells = [c.strip() for c in line.split(",")]
                rows.append(cells)
        self._run_tool("excel_create_table", {
            "filename": self.xl_filename.text(),
            "headers": headers,
            "rows": rows,
        })

    def _direct_excel_insert(self):
        val = self.xl_ins_val.text()
        try:
            val = int(val)
        except ValueError:
            try:
                val = float(val)
            except ValueError:
                pass
        self._run_tool("excel_insert_value", {
            "filename": self.xl_ins_file.text(),
            "row": self.xl_ins_row.value(),
            "column": self.xl_ins_col.value(),
            "value": val,
        })

    def _direct_ppt_create(self):
        self._run_tool("ppt_create_presentation", {
            "filename": self.ppt_filename.text(),
            "title": self.ppt_title.text(),
            "subtitle": self.ppt_subtitle.text(),
            "company": self.ppt_company.text(),
            "date": self.ppt_date.text(),
        })

    def _direct_ppt_add_slide(self):
        bullets = [l.strip() for l in self.ppt_add_content.toPlainText().split("\n") if l.strip()]
        self._run_tool("ppt_add_slide", {
            "filename": self.ppt_add_file.text(),
            "title": self.ppt_add_title.text(),
            "content": bullets,
        })

    def _direct_pdf_extract_text(self):
        self._run_tool("pdf_extract_text", {
            "pdf_path": self.pdf_txt_src.text(),
            "page_number": self.pdf_txt_page.value(),
        })

    def _direct_word_to_pdf(self):
        self._run_tool("pdf_word_to_pdf", {
            "docx_path": self.pdf_docx.text(),
            "pdf_path": self.pdf_out_pdf.text(),
        })

    def _direct_ocr_extract_text(self):
        image_path = self.ocr_image_path.text().strip()
        if not image_path:
            self.ocr_result.setPlainText("Please select or paste an image first.")
            return

        result = execute_tool("ocr_image_to_text", {
            "image_path": image_path,
            "language": self.ocr_lang.currentText(),
            "save_to": self.ocr_save_txt.text().strip(),
        }, raw=True)

        self.ocr_result.setPlainText(result.get("text", result.get("message", "")))
        self._append_chat(tool_bubble(result.get("message", "OCR completed")))
        self._status(result.get("message", "OCR completed")[:80])

    def _direct_ocr_to_word(self):
        image_path = self.ocr_image_path.text().strip()
        if not image_path:
            self.ocr_result.setPlainText("Please select or paste an image first.")
            return

        output_docx = self.ocr_save_docx.text().strip()
        if not output_docx:
            base_name = os.path.splitext(os.path.basename(image_path))[0] or "ocr_output"
            output_docx = os.path.join(os.getcwd(), "output", f"{base_name}_ocr.docx")
            self.ocr_save_docx.setText(output_docx)

        result = execute_tool("ocr_image_to_word", {
            "image_path": image_path,
            "output_docx": output_docx,
            "language": self.ocr_lang.currentText(),
        }, raw=True)

        self.ocr_result.setPlainText(result.get("text", result.get("message", "")))
        self._append_chat(tool_bubble(result.get("message", "OCR Word export completed")))
        self._status(result.get("message", "OCR Word export completed")[:80])

    def _save_ocr_result(self):
        text = self.ocr_result.toPlainText().strip()
        if not text:
            self.ocr_result.setPlainText("No OCR result available to save yet.")
            return

        save_path = self.ocr_save_txt.text().strip()
        if not save_path:
            save_path, _ = QFileDialog.getSaveFileName(self, "Save OCR Text", "", "Text Files (*.txt)")
            if not save_path:
                return
            self.ocr_save_txt.setText(save_path)

        try:
            os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
            with open(save_path, "w", encoding="utf-8") as handle:
                handle.write(text)
            message = f"OCR result saved to: {save_path}"
            self._append_chat(tool_bubble(message))
            self._status(message[:80])
        except Exception as exc:
            self.ocr_result.setPlainText(f"Failed to save OCR result: {exc}")
