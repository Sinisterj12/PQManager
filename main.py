"""
Printer Queue Manager
Main entry point for the application
"""

import logging
import os
from src.main import Application

def main():
    app = Application()
    app.run()

if __name__ == "__main__":
    main()
