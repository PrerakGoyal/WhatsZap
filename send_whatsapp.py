"""
WhatsApp Bulk Messenger - Simple Version
Just put your contacts in contacts.xlsx and run this script!
"""

import time
import os
import shutil
import traceback
import tempfile
import pandas as pd
from selenium.common.exceptions import SessionNotCreatedException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def send_whatsapp_messages():
    """Main function to send WhatsApp messages"""
    
    print("=" * 50)
    print("WhatsApp Bulk Messenger")
    print("=" * 50)
    
    # Step 1: Read contacts from Excel
    print("\n📊 Reading contacts from contacts.xlsx...")
    try:
        df = pd.read_excel('contacts.xlsx')
        
        # Check if required columns exist
        if 'phone' not in df.columns or 'message' not in df.columns:
            print("❌ Error: Excel file must have 'phone' and 'message' columns!")
            print("   Example:")
            print("   phone        | message")
            print("   919876543210 | Hello, this is a test message")
            return
        
        # Remove empty rows
        df = df.dropna(subset=['phone', 'message'])
        contacts = df.to_dict('records')
        
        print(f"✅ Found {len(contacts)} contacts to message")
        
        if len(contacts) == 0:
            print("❌ No contacts found in the Excel file!")
            return
            
    except FileNotFoundError:
        print("❌ Error: contacts.xlsx not found!")
        print("   Please create an Excel file named 'contacts.xlsx' with columns:")
        print("   phone        | message")
        print("   919876543210 | Hello, this is a test message")
        return
    except Exception as e:
        print(f"❌ Error reading Excel file: {str(e)}")
        return
    
    # Step 2: Setup Chrome browser
    print("\n🌐 Setting up Chrome browser...")
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--user-data-dir=./whatsapp_session")  # Save login
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        # Attempt to locate Chrome binary on Windows
        chrome_path = None
        if os.name == 'nt':
            possible = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
            ]
            for p in possible:
                if p and os.path.exists(p):
                    chrome_path = p
                    break
            if not chrome_path:
                found = shutil.which("chrome") or shutil.which("chrome.exe")
                if found:
                    chrome_path = found
        if chrome_path:
            options.binary_location = chrome_path
            print(f"Using Chrome binary: {chrome_path}")
        else:
            print("⚠️  Could not find Chrome executable automatically. Ensure Chrome is installed.")

        # Auto-install ChromeDriver and start browser
        service = Service(ChromeDriverManager().install())

        def try_start_driver(opts):
            return webdriver.Chrome(service=service, options=opts)

        try:
            driver = try_start_driver(options)
            wait = WebDriverWait(driver, 30)

        except Exception as e:
            msg = str(e)
            print(f"❌ Error starting Chrome (first attempt): {msg}")
            traceback.print_exc()

            # Common cause: existing Chrome profile or port issue. Retry with a fresh temporary profile
            if "DevToolsActivePort" in msg or isinstance(e, SessionNotCreatedException) or "session not created" in msg:
                print("ℹ️  Detected DevToolsActivePort or session error — retrying with a temporary profile and safer flags...")
                try:
                    temp_profile = tempfile.mkdtemp(prefix="wa_profile_")
                    alt_opts = webdriver.ChromeOptions()
                    alt_opts.add_argument(f"--user-data-dir={temp_profile}")
                    alt_opts.add_argument("--remote-debugging-port=9222")
                    alt_opts.add_argument("--disable-gpu")
                    alt_opts.add_argument("--no-sandbox")
                    alt_opts.add_argument("--disable-dev-shm-usage")
                    alt_opts.add_experimental_option("excludeSwitches", ["enable-logging"])
                    if chrome_path:
                        alt_opts.binary_location = chrome_path

                    driver = try_start_driver(alt_opts)
                    wait = WebDriverWait(driver, 30)
                    print("✅ Browser started using a temporary profile (your saved session won't be used this run).")

                except Exception as e2:
                    print(f"❌ Retry failed: {str(e2)}")
                    traceback.print_exc()
                    return
            else:
                print("❌ Unexpected error starting Chrome")
                return
        
        print("✅ Browser ready")
        
    except Exception as e:
        print(f"❌ Error setting up browser: {str(e)}")
        traceback.print_exc()
        print("   Make sure Chrome is installed on your computer and is compatible with the installed chromedriver")
        return
    
    # Step 3: Open WhatsApp Web
    print("\n💬 Opening WhatsApp Web...")
    driver.get("https://web.whatsapp.com")
    
    print("\n⚠️  IMPORTANT:")
    print("   If this is your first time:")
    print("   1. Scan the QR code with your phone")
    print("   2. Wait for WhatsApp to load completely")
    print("   3. You'll see your chats appear")
    print("\n   Next time, you won't need to scan again!")
    print("\n   Waiting for WhatsApp to load...")
    
    # Wait for WhatsApp to load
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
        ))
        print("✅ WhatsApp loaded successfully!")
        time.sleep(3)
        
    except Exception as e:
        print("❌ Timeout waiting for WhatsApp. Please try again.")
        driver.quit()
        return
    
    # Step 4: Send messages
    print("\n📤 Starting to send messages...")
    print("-" * 50)
    
    successful = 0
    failed = 0
    
    for i, contact in enumerate(contacts, 1):
        phone = str(contact['phone']).strip().replace('+', '').replace(' ', '')
        message = str(contact['message']).strip()
        
        print(f"\n[{i}/{len(contacts)}] Sending to {phone}...")
        
        try:
            # Open chat using WhatsApp's click-to-chat
            driver.get(f"https://web.whatsapp.com/send?phone={phone}")
            time.sleep(4)  # Wait for chat to open
            
            # Find the message input box
            try:
                chat_box = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
                ))
            except:
                # Try alternative selector
                chat_box = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@contenteditable="true"][@data-lexical-editor="true"]')
                ))
            
            # Type and send message
            chat_box.click()
            time.sleep(1)
            
            # Type message line by line (handles multi-line messages)
            lines = message.split('\n')
            for line in lines[:-1]:
                chat_box.send_keys(line)
                chat_box.send_keys(Keys.SHIFT, Keys.ENTER)
            chat_box.send_keys(lines[-1])
            
            time.sleep(1)
            chat_box.send_keys(Keys.ENTER)
            
            print(f"   ✅ Sent successfully!")
            successful += 1
            time.sleep(3)  # Wait before next message
            
        except Exception as e:
            print(f"   ❌ Failed: {str(e)}")
            failed += 1
            time.sleep(2)
    
    # Step 5: Summary
    print("\n" + "=" * 50)
    print("📊 SUMMARY")
    print("=" * 50)
    print(f"✅ Successful: {successful}")
    print(f"❌ Failed: {failed}")
    print(f"📱 Total: {len(contacts)}")
    print("\n✨ Done! You can close the browser now.")
    
    # Keep browser open for 10 seconds
    time.sleep(10)
    driver.quit()


if __name__ == "__main__":
    try:
        send_whatsapp_messages()
    except KeyboardInterrupt:
        print("\n\n⚠️  Stopped by user")
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
    
    input("\nPress Enter to exit...")
