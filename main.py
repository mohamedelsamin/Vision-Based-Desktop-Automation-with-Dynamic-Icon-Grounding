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
import subprocess
import urllib.request
import urllib.error
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

# UTILITY FUNCTIONS 


def fetch_posts_urllib():
    """Fetch posts using urllib (built-in, no external dependencies)."""
    print("Attempting to fetch posts using urllib...")
    try:
        req = urllib.request.Request(
            POSTS_API,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        )
        with urllib.request.urlopen(req, timeout=15) as response:
            data = response.read().decode('utf-8')
            posts = json.loads(data)[:MAX_POSTS]
            print(f"Successfully fetched {len(posts)} posts using urllib.")
            return posts
    except urllib.error.URLError as e:
        print(f"Error with urllib: {e}")
        return []
    except Exception as e:
        print(f"Error fetching with urllib: {e}")
        return []

def fetch_posts_requests_improved():
    """Fetch posts using requests with improved error handling and SSL options."""
    print("Attempting to fetch posts using requests (improved method)...")
    try:
        import ssl
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        session = requests.Session()
        session.verify = False  # Disable SSL verification if needed
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive'
        }
        
        response = session.get(POSTS_API, headers=headers, timeout=20, allow_redirects=True)
        response.raise_for_status()
        posts = response.json()[:MAX_POSTS]
        print(f"Successfully fetched {len(posts)} posts using requests (improved).")
        return posts
    except Exception as e:
        print(f"Error with improved requests method: {e}")
        return []

def fetch_posts_via_default_browser():
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

def fetch_posts_via_edge():
    """Fetch posts by opening Microsoft Edge and copying the JSON response."""
    print("Opening Microsoft Edge to fetch posts from API...")
    try:
        # Clear clipboard first
        pyperclip.copy("")
        
        # Open Edge via Windows search
        pyautogui.hotkey('win')
        time.sleep(1.5)
        pyautogui.write('edge')
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(5)
        
        # Navigate to the API URL
        pyautogui.hotkey('ctrl', 'l')
        time.sleep(1.5)
        pyautogui.write(POSTS_API)
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(8)
        
        # Click and copy
        pyautogui.click(400, 300)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(3)
        
        data = pyperclip.paste()
        
        if not data or not data.strip().startswith('['):
            print("Retrying...")
            time.sleep(2)
            pyautogui.click(400, 300)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1.5)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(3)
            data = pyperclip.paste()
        
        if not data or not data.strip().startswith('['):
            print(f"Error: Clipboard doesn't contain JSON.")
            pyautogui.hotkey('ctrl', 'w')
            return []
        
        posts = json.loads(data)
        print(f"Successfully fetched {len(posts)} posts from Edge.")
        
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(1)
        
        return posts[:MAX_POSTS]
    except Exception as e:
        print(f"Error fetching from Edge: {e}")
        try:
            pyautogui.hotkey('ctrl', 'w')
        except:
            pass
        return []

def fetch_posts_via_subprocess():
    """Open the API URL using subprocess (Windows start command)."""
    print("Opening API URL using subprocess...")
    try:
        # Clear clipboard
        pyperclip.copy("")
        
        # Open URL using Windows start command
        subprocess.Popen(['start', POSTS_API], shell=True)
        time.sleep(6)
        
        # Copy JSON from browser
        pyautogui.click(400, 300)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(3)
        
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
            print(f"Error: Clipboard doesn't contain JSON.")
            pyautogui.hotkey('alt', 'f4')
            return []
        
        posts = json.loads(data)
        print(f"Successfully fetched {len(posts)} posts via subprocess.")
        
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)
        
        return posts[:MAX_POSTS]
    except Exception as e:
        print(f"Error with subprocess method: {e}")
        try:
            pyautogui.hotkey('alt', 'f4')
        except:
            pass
        return []

def fetch_posts():
    """Fetch the first 10 posts from the API with multiple fallback methods."""
    print("=" * 60)
    print("Fetching posts from API - trying multiple methods...")
    print("=" * 60)
    
    # Method 1: Try requests library (original)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    for attempt in range(2):
        try:
            print(f"\n[Method 1] Attempting requests library (attempt {attempt + 1}/2)...")
            response = requests.get(POSTS_API, headers=headers, timeout=15, verify=False)
            response.raise_for_status()
            posts = response.json()[:MAX_POSTS]
            print(f"[SUCCESS] Fetched {len(posts)} posts using requests library.")
            return posts
        except requests.exceptions.RequestException as e:
            print(f"[FAILED] Requests library error: {e}")
            if attempt < 1:
                print("Retrying in 2 seconds...")
                time.sleep(2)
    
    # Method 2: Try improved requests method
    print(f"\n[Method 2] Trying improved requests method...")
    posts = fetch_posts_requests_improved()
    if posts:
        return posts
    
    # Method 3: Try urllib (built-in, no dependencies)
    print(f"\n[Method 3] Trying urllib (built-in library)...")
    posts = fetch_posts_urllib()
    if posts:
        return posts
    
    # Method 4: Try default browser
    print(f"\n[Method 4] Trying default browser...")
    posts = fetch_posts_via_default_browser()
    if posts:
        return posts
    
    # Method 5: Try Microsoft Edge
    print(f"\n[Method 5] Trying Microsoft Edge...")
    posts = fetch_posts_via_edge()
    if posts:
        return posts
    
    # Method 6: Try subprocess
    print(f"\n[Method 6] Trying subprocess method...")
    posts = fetch_posts_via_subprocess()
    if posts:
        return posts
    
    print("\n[FAILED] All methods exhausted. Could not fetch posts.")
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
    
    # Save file - automatically replace existing files
    filename = f"post_{post['id']}.txt"
    full_path = os.path.join(OUTPUT_DIR, filename)
    
    # Delete existing file if it exists (automatic replacement)
    if os.path.exists(full_path):
        try:
            # Try to delete the file
            os.remove(full_path)
            # Wait a moment and verify deletion
            time.sleep(0.1)
            # Retry deletion if file still exists (handles file locks)
            if os.path.exists(full_path):
                time.sleep(0.2)
                os.remove(full_path)
            
            # Verify file is actually deleted
            if os.path.exists(full_path):
                raise Exception("File still exists after deletion attempt")
            
            print(f"Deleted existing file: {filename}")
        except Exception as e:
            print(f"Warning: Could not delete existing file {filename}: {e}")
            print(f"Attempting to overwrite existing file instead...")
    
    # Ensure file doesn't exist before creating new one
    if os.path.exists(full_path):
        # Last attempt: try to remove with different method
        try:
            os.chmod(full_path, stat.S_IWRITE)  # Make file writable
            os.remove(full_path)
            time.sleep(0.1)
        except:
            pass
    
    # Method 1: Use Notepad save dialog (primary method)
    try:
        # Open save dialog
        pyautogui.hotkey('ctrl', 's')
        time.sleep(1)
        
        # Type the full path
        pyautogui.write(full_path)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
    
      
        
        # Verify file was saved
        if os.path.exists(full_path):
            print(f"Post {post['id']} saved as {filename} (replaced old file if present)")
        else:
            raise Exception("File was not saved via Notepad dialog")
    except Exception as e:
        print(f"Error saving via Notepad dialog: {e}. Using direct file write method...")
        
    
    # Close Notepad
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(0.5)

def fallback_open_notepad_via_search(main_window_title="Untitled - Notepad"):
    """Open Notepad using Windows search as a fallback and close popups."""
    print("Fallback: Opening Notepad via Windows search...")
    pyautogui.hotkey('win')
    pyautogui.write("Notepad")
    time.sleep(0.5)
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
