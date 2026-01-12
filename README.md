# TJM-Task
# Vision-Based Desktop Automation Demo

## Overview
This project demonstrates a Python desktop automation system that:
- Fetches data from an external API
- Detects desktop icons using computer vision
- Automates Notepad interaction
- Saves structured content automatically

## Key Features
- Multi-method API fetching with fallback handling
- Identify the icon regardless of its position
- Return the center coordinates (x, y) for clicking
- Handle cases where the icon might be in different locations
- Vision-based icon detection (OpenCV)
- Desktop automation using PyAutoGUI
- Robust fallback logic
- Handle existing files in target directory
- Detect multiple desktop icons and select the correct one
- Support both light and dark desktop themes


## Tech Stack
- Python
- OpenCV
- PyAutoGUI
- mss
- Requests / urllib
