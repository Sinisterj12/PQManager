"""
Bootstrap script for PQ Manager
"""
import os
import sys
import logging

# Configure logging to write to both file and console
def setup_logging():
    # Create logs directory if it doesn't exist
    if not os.path.exists('logs'):
        os.makedirs('logs')

    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('logs/printer_queue.log'),
            logging.StreamHandler(sys.stdout)  # Add console output
        ]
    )

# Add the project root directory to Python path
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Initialize environment
os.environ['PYTHONPATH'] = project_root

def main():
    setup_logging()
    logging.info("Starting PQ Manager...")
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Project root: {project_root}")
    
    try:
        # Import and run the main application
        from src.main import Application
        app = Application()
        app.run()
    except Exception as e:
        logging.error(f"Error starting application: {e}", exc_info=True)
        input("Press Enter to exit...")  # Keep console open on error

if __name__ == "__main__":
    main()
