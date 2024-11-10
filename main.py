import os
import sys

# Add the project root directory to Python path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

# Now import and run the actual application
from src.main import *

if __name__ == "__main__":
    root.mainloop()
