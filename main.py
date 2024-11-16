"""
Printer Queue Manager
Main entry point for the application
"""

import os
import sys

# Add the project root directory to Python path
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from src.main import Application

def main():
    app = Application()
    app.run()

if __name__ == "__main__":
    main()
