import os
import time
import stat
import pyautogui
import pygetwindow as gw
import cv2
import numpy as np
import requests
import json
import pyperclip
import webbrowser
import mss


# Import configuration constants
from config import (
    ICON_PATH,
    OUTPUT_DIR,
    ANNOTATED_DIR,
    POSTS_API,
    MAX_POSTS,
    RETRY_ATTEMPTS,
    RETRY_DELAY
)


def fallback_fetch_posts_via_default_browser():
    """Open the API URL in default browser and copy JSON from clipboard."""
    print("Opening API URL in default browser...")
    try:
        # Clear clipboard first
        pyperclip.copy("")
        
        # Open URL in default browser
        webbrowser.open(POSTS_API)
        time.sleep(6)  # Wait for browser to open and load
        
        # Click on the page to ensure it's focused
        pyautogui.click(400, 300)
        time.sleep(1)
        
        # Select all and copy
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(3)
        
        # Get data from clipboard
        data = pyperclip.paste()
        
        if not data or not data.strip().startswith('['):
            print("Retrying clipboard copy...")
            time.sleep(1)
            pyautogui.click(400, 300)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1.5)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(3)
            data = pyperclip.paste()
        
        if not data or not data.strip().startswith('['):
            print(f"Error: Clipboard doesn't contain JSON. Content: {data[:200]}...")
            # Minimize browser instead of closing
            try:
                chrome_windows = gw.getWindowsWithTitle("Chrome")
                if chrome_windows:
                    chrome_windows[0].minimize()
                    print("Browser minimized.")
            except:
                pyautogui.hotkey('win', 'down')  # Alternative minimize method
            return []
        
        posts = json.loads(data)
        print(f"Successfully fetched {len(posts)} posts from default browser.")
        
        # Minimize browser instead of closing
        try:
            # Try to find Chrome window by title
            chrome_windows = gw.getWindowsWithTitle("Chrome")
            if not chrome_windows:
                # Try alternative browser names
                chrome_windows = gw.getWindowsWithTitle("Google Chrome")
            if not chrome_windows:
                chrome_windows = gw.getWindowsWithTitle("chromium")
            
            if chrome_windows:
                chrome_windows[0].minimize()
                print("Browser minimized successfully.")
            else:
                # Fallback: Use keyboard shortcut to minimize
                pyautogui.hotkey('win', 'down')
                print("Browser minimized using keyboard shortcut.")
        except Exception as e:
            print(f"Could not minimize browser: {e}. Using keyboard shortcut.")
            pyautogui.hotkey('win', 'down')  # Alternative minimize method
        
        time.sleep(0.5)
        
        return posts[:MAX_POSTS]
    except Exception as e:
        print(f"Error fetching from default browser: {e}")
        try:
            # Minimize browser instead of closing
            chrome_windows = gw.getWindowsWithTitle("Chrome")
            if not chrome_windows:
                chrome_windows = gw.getWindowsWithTitle("Google Chrome")
            if chrome_windows:
                chrome_windows[0].minimize()
            else:
                pyautogui.hotkey('win', 'down')
        except:
            pass
        return []


def fetch_posts():
    """Fetch posts from the API - try to open/fetch link directly, fallback to browser method."""
    try:
        response = requests.get(POSTS_API)
        response.raise_for_status()
        return response.json()[:MAX_POSTS]
    except:
        print("API unavailable, opening Chrome to fetch posts.")
        print("=" * 60)
        posts = fallback_fetch_posts_via_default_browser()
        if posts:
            return posts
        print("\n[FAILED] Could not fetch posts using any method.")
        return []

def close_unexpected_popups(main_window_title):
    
    windows = gw.getAllTitles()
    for w in windows:
        if w and main_window_title not in w:
            try:
                win = gw.getWindowsWithTitle(w)[0]
                win.activate()
                time.sleep(0.3)
                pyautogui.press('esc')
                time.sleep(0.2)
                pyautogui.hotkey('alt', 'f4')
                time.sleep(0.2)
            except Exception as e:
                print(f"Could not close window '{w}': {e}")

# NOTEPAD FUNCTIONS 

def open_notepad(screenshot_index=0, scales=[0.5, 0.75, 1.0, 1.25, 1.5, 2.0], threshold=0.5):

    # Minimize all windows
    pyautogui.hotkey('win', 'd')
    time.sleep(1)

    # Take a desktop screenshot using mss
    screenshot_path = os.path.join(ANNOTATED_DIR, f"icon_detect{screenshot_index}.png")
    with mss.mss() as sct:
        # Get the primary monitor
        monitor = sct.monitors[1]
        screenshot = sct.grab(monitor)
        # Convert to numpy array and save
        img = np.array(screenshot)
        # mss returns BGRA, convert to BGR for OpenCV
        desktop_color = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        cv2.imwrite(screenshot_path, desktop_color)
    
    desktop_gray = cv2.cvtColor(desktop_color, cv2.COLOR_BGR2GRAY)

    # Edge detection to highlight desktop icons
    desktop_edges = cv2.Canny(desktop_gray, 50, 150)
    contours, _ = cv2.findContours(desktop_edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > 20 and h > 20:
            cv2.rectangle(desktop_color, (x, y), (x + w, y + h), (0, 0, 255), 2)

    # Load target icon and apply edge detection
    target_img = cv2.imread(ICON_PATH, cv2.IMREAD_GRAYSCALE)
    if target_img is None:
        print("Icon image not found.")
        return False
    target_edges = cv2.Canny(target_img, 50, 150)

    # Multi-scale template matching to find the icon
    target_found = False
    for scale in scales:
        resized_target = cv2.resize(target_edges, (0, 0), fx=scale, fy=scale)
        res = cv2.matchTemplate(desktop_edges, resized_target, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)

        if max_val >= threshold:
            x, y = max_loc
            h, w = resized_target.shape
            cv2.rectangle(desktop_color, (x, y), (x + w, y + h), (0, 255, 0), 3)

            center_x = x + w // 2
            center_y = y + h // 2
            pyautogui.moveTo(center_x, center_y)
            pyautogui.doubleClick()
            print(f"Notepad icon found and opened at ({center_x}, {center_y})")
            target_found = True
            break

    # Save annotated screenshot
    cv2.imwrite(screenshot_path, desktop_color)
    print(f"Annotated desktop saved as {screenshot_path}")

    return target_found

def wait_for_notepad(timeout=10):

    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = gw.getWindowsWithTitle("Untitled - Notepad")
        if windows:
            try:
                windows[0].activate()
                time.sleep(0.5)
            except:
                pass
            if windows[0].isActive:
                return windows[0]
        time.sleep(0.5)
    return None

def type_and_save_post(post):
    """Type the post content and save it as a file."""
    # Format the content
    content = f"Title: {post['title']}\n\n{post['body']}\n\n"
    
    # Type each line
    for line in content.splitlines():
        pyautogui.write(line, interval=0.03)
        pyautogui.press('enter')
    time.sleep(0.5)
    
    # Save file - keep old file and automatically handle replacement popup
    filename = f"post_{post['id']}.txt"
    full_path = os.path.join(OUTPUT_DIR, filename)
    
    # Check if file exists to inform user
    if os.path.exists(full_path):
        print(f"File {filename} already exists. Replacement popup will be handled automatically.")
    
    # Use Notepad save dialog - let Windows show replacement popup if file exists
    # Open save dialog
    pyautogui.hotkey('ctrl', 's')
    time.sleep(1)
    
    # Type the full path
    pyautogui.write(full_path)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.8)  # Wait for replacement dialog to appear if file exists
    
    # Automatically handle Windows "File already exists" confirmation dialog
    # The dialog asks: "Do you want to replace it?" with Yes/No buttons
    # We automatically press 'y' (Yes) or Enter to confirm replacement
    try:
        # Wait a bit more to ensure dialog has appeared (if file exists)
        time.sleep(0.3)
        # Press 'y' to confirm replacement (works for Yes button)
        pyautogui.press('y')
        time.sleep(0.2)
        print("Automatically confirmed file replacement")
    except:
        # If 'y' doesn't work, try Enter key (some dialogs use Enter as default)
        try:
            time.sleep(0.2)
            pyautogui.press('enter')
        except:
            pass
    
    time.sleep(0.5)
    
    # Verify file was saved
    if os.path.exists(full_path):
        print(f"Post {post['id']} saved as {filename} (replaced old file automatically)")
    else:
        print(f"Warning: Could not verify if file {filename} was saved")
    
    # Close Notepad
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(0.5)

def fallback_open_notepad_via_search(main_window_title="Untitled - Notepad"):
    """Open Notepad using Run dialog (Win+R) with Notepad path as a fallback."""
    print("Fallback: Opening Notepad via Run dialog (Win+R)...")
    
    # Open Run dialog with Win+R
    pyautogui.hotkey('win', 'r')
    time.sleep(0.5)
    
    # Type Notepad path - using the full path to notepad.exe
    # Common paths: C:\Windows\System32\notepad.exe or just notepad.exe
    notepad_path = "notepad.exe"
    pyautogui.write(notepad_path)
    time.sleep(0.5)
    
    # Press Enter to execute
    pyautogui.press('enter')
    time.sleep(1.5)  # wait for Notepad to open

    # Close any unexpected popups
    close_unexpected_popups(main_window_title)

# MAIN EXECUTION 

if __name__ == "__main__":
    main_window_title = "Untitled - Notepad"

    # Step 1: Fetch posts from API
    print("Fetching posts from API...")
    posts = fetch_posts()
    if not posts:
        print("No posts available. Exiting.")
        exit(1)
    
    # Step 2: Process each post
    for idx, post in enumerate(posts):
        print(f"\nProcessing post {post['id']} ({idx + 1}/{len(posts)})...")
        
        # Attempt to open Notepad
        success = False
        for attempt in range(RETRY_ATTEMPTS):
            if open_notepad(screenshot_index=idx):
                notepad_window = wait_for_notepad()
                if notepad_window:
                    success = True
                    print("Notepad opened successfully!")
                    break
            print(f"Attempt {attempt + 1} failed to open Notepad. Retrying in {RETRY_DELAY} sec...")
            time.sleep(RETRY_DELAY)

        # If still not successful, use fallback
        if not success:
            print("Icon detection failed. Using fallback method...")
            fallback_open_notepad_via_search(main_window_title)
            notepad_window = wait_for_notepad()
            if notepad_window:
                success = True
                print("Notepad opened successfully using fallback method!")

        if not success:
            print(f"Failed to open Notepad for post {post['id']}. Skipping...")
            continue

        # Type and save the post
        type_and_save_post(post)
        time.sleep(1)
    
    print(f"\nAll {len(posts)} posts processed successfully!")
