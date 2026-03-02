"""
main.py
Entry point for tEppy's Data Entry App.
"""

from ttkbootstrap import Window
from app import DynamicExcelApp, APP_TITLE


def main():
    root = Window(title=APP_TITLE, themename="cosmo")
    DynamicExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
