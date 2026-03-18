
import sys
import os
from dotenv import load_dotenv
from PySide6.QtWidgets import QApplication
from src.ui.main_window import MainWindow


def main():
    load_dotenv(override=True)
    api_key = os.getenv("OPENAI_API_KEY", "")

    app = QApplication(sys.argv)
    app.setApplicationName("AI Agent - Office Automation")
    app.setStyle("Fusion")

    window = MainWindow(api_key=api_key)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
