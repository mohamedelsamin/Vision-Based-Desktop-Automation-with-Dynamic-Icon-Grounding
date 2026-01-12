"""
Configuration constants for tjm-project
"""

import os

# Icon Detection
ICON_PATH = "notepad.png"

# Output Directories
OUTPUT_DIR = r"C:\Users\Mohamed\OneDrive\Desktop\tjm-project"
ANNOTATED_DIR = os.path.join(OUTPUT_DIR, "annotated_screenshot")

# API Configuration
POSTS_API = "https://jsonplaceholder.typicode.com/posts"

# Processing Configuration
MAX_POSTS = 1
RETRY_ATTEMPTS = 3
RETRY_DELAY = 1  # seconds

# Create required directories on import
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ANNOTATED_DIR, exist_ok=True)
