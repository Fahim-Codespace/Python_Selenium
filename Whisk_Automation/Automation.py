from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import subprocess



# Setup Chrome options
user_data_dir = r"C:\Users\fahim siam\AppData\Local\Google\Chrome\User Data2"
profile_dir = "SeleniumProfile"

chrome_options = Options()
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
chrome_options.add_argument(f"--profile-directory={profile_dir}")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--no-first-run")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--remote-debugging-port=9222")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
chrome_options.add_experimental_option('useAutomationExtension', False)

print(f"Using profile: {profile_dir}")

# Initialize driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 15)
actions = ActionChains(driver)

def wait_for_manual_signin():
    """Wait for user to manually sign in to Whisk"""
    print("\n" + "="*60)
    print("MANUAL SIGN-IN REQUIRED")
    print("="*60)
    print("Please sign in to Whisk manually in the opened browser.")
    print("Steps:")
    print("1. The Whisk tab should be open and visible")
    print("2. Sign in with your Google account if prompted")
    print("3. Wait until you see the Whisk image generation interface")
    print("4. Make sure the input field is visible and ready")
    print("5. Come back here and press Enter to continue automation")
    print("="*60)
    
    input("Press Enter AFTER you have successfully signed in to Whisk...")
    print("Resuming automation...")
    time.sleep(2)

def check_whisk_ready():
    """Check if Whisk is ready for automation"""
    print("Checking if Whisk is ready...")
    
    # Check for input field
    input_selectors = [
        "//input[contains(@placeholder, 'Describe your idea')]",
        "//input[contains(@placeholder, 'prompt ideas')]",
        "//textarea",
        "//input[@type='text']"
    ]
    
    for selector in input_selectors:
        try:
            elements = driver.find_elements(By.XPATH, selector)
            if elements and elements[0].is_displayed():
                print("✓ Whisk input field found - ready for automation")
                return True
        except:
            continue
    
    print("✗ Whisk input field not found")
    return False

try:
    # Open Google sheets in First tab
    driver.get("https://docs.google.com/spreadsheets/d/1sVhAn5qrHx1xMAzKoBTSRvVEptp_mxoTdEi0UZMXjQk/edit?gid=0#gid=0")
    print("Google Sheets opened")
    time.sleep(3)
    
    # Open Whisk in new tab
    driver.execute_script("window.open('https://labs.google/fx/tools/whisk/project','_blank');")
    print("Whisk tab opened")
    
    # Store window handles
    sheets_tab = driver.window_handles[0]
    whisk_tab = driver.window_handles[1]
    
    print("Both tabs opened successfully!")
    
    # Switch to Whisk tab and wait for manual sign-in
    driver.switch_to.window(whisk_tab)
    print("Switched to Whisk tab")
    
    # Wait for page to load
    time.sleep(5)
    
    # Check if we're on sign-in page
    signin_indicators = [
        "//*[contains(text(), 'Sign in')]"
        
    ]
    
    needs_signin = False
    for indicator in signin_indicators:
        try:
            if driver.find_elements(By.XPATH, indicator):
                needs_signin = True
                break
        except:
            continue
    
    if needs_signin:
        print("Whisk sign-in page detected.")
        wait_for_manual_signin()
    else:
        print("Whisk appears to be already signed in.")
        if not check_whisk_ready():
            print("Whisk doesn't seem ready. Waiting for manual verification...")
            wait_for_manual_signin()
    
    # Final check before starting automation
    if not check_whisk_ready():
        print("Whisk still not ready. Please check the browser and press Enter when ready...")
        input("Press Enter when Whisk is ready...")
    
    print("Starting automation in 3 seconds...")
    time.sleep(3)
    
    # Switch back to Google sheets tab to start automation
    driver.switch_to.window(sheets_tab)
    time.sleep(2)
    
    # Counter for cells
    cell_count = 1
    empty_cells_count = 0
    max_empty_cells = 3
    
    # Initial selection of A1
    print("Selecting cell A1 in Google Sheets...")
    actions.key_down(Keys.CONTROL).send_keys(Keys.HOME).key_up(Keys.CONTROL).perform()
    time.sleep(1)
    actions.key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
    time.sleep(0.25)
    
    print("Automation starting NOW!")
    print("="*50)
    
    while True:
        try:
            print(f"\n--- Processing cell {cell_count} ---")
            
            # Switch to Google Sheets tab
            driver.switch_to.window(sheets_tab)
            time.sleep(0.5)
            
            # Copy the current cell content (Ctrl + c)
            actions.key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
            time.sleep(0.50)
            
            # Switch to whisk tab
            driver.switch_to.window(whisk_tab)
            time.sleep(0.5)
            
            # Find the text input box
            text_input = None
            input_selectors = [
                "//input[contains(@placeholder, 'Describe your idea')]",
                "//input[contains(@placeholder, 'prompt ideas')]",
                "//textarea",
                "//input[@type='text']"
            ]
            
            for selector in input_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    if elements:
                        text_input = elements[0]
                        break
                except:
                    continue
            
            if text_input:
                # Click on the input field
                text_input.click()
                time.sleep(0.25)
                
                # Clear any existing text (Ctrl + A, Backspace)
                text_input.send_keys(Keys.CONTROL + 'a')
                time.sleep(0.25)
                text_input.send_keys(Keys.DELETE)
                time.sleep(0.25)
                
                # Paste the copied content (Ctrl + V)
                text_input.send_keys(Keys.CONTROL + 'v')
                time.sleep(0.5)
                
                # Check if we pasted something (not empty)
                pasted_content = text_input.get_attribute('value')
                if not pasted_content or pasted_content.strip() == "":
                    print(f"Cell A{cell_count} is empty. Skipping submission.")
                    empty_cells_count += 1
                    if empty_cells_count >= max_empty_cells:
                        print(f"Encountered {max_empty_cells} consecutive empty cells. Stopping automation.")
                        break
                else:
                    print(f"Pasted content from cell A{cell_count}: {pasted_content[:50]}...")
                    empty_cells_count = 0
                    
                    # Submit the input (Enter)
                    text_input.send_keys(Keys.ENTER)
                    print("Submitted the input.")
                    
                    # Wait for generation
                    print("Waiting for image generation...")
                    time.sleep(14)
                    
                    # Clear the text box for next prompt
                    try:
                        text_input.click()
                        text_input.send_keys(Keys.CONTROL + 'a')
                        time.sleep(0.25)
                        text_input.send_keys(Keys.DELETE)
                        print("Cleared text box.")
                    except:
                        print("Could not clear the text box.")
                        
            else:
                print("ERROR: Could not find the text input box on Whisk!")
                print("Automation paused. Please fix the issue and press Enter to continue...")
                input("Press Enter after fixing Whisk...")
                continue
            
            # Move to next cell in Google Sheets (Down Arrow)
            driver.switch_to.window(sheets_tab)
            actions.send_keys(Keys.DOWN).perform()
            time.sleep(0.5)
            cell_count += 1
            
        except KeyboardInterrupt:
            print("\nAutomation interrupted by user.")
            break
        except Exception as e:
            print(f"Unexpected error: {e}")
            print("Trying to continue with next cell...")
            cell_count += 1
            continue
    
    print("\n" + "="*50)
    print("Automation completed!")
    print(f"Total cells processed: {cell_count - 1}")
    
except Exception as e:
    print(f"Critical error: {e}")

finally:
    input("\nPress Enter to close the browser...")
    driver.quit()
    print("Browser closed.")