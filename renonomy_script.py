from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
import webbrowser

print("Script started...")

try:
    # Set up the Safari driver
    driver = webdriver.Safari()

    # Maximize the window to full screen
    driver.maximize_window()

    # Open the Reonomy login page
    renonomy_login_link = "https://app.reonomy.com/!/auth/login"
    driver.get(renonomy_login_link)
    print("Opened Reonomy login page.")

    # Wait for the login page to load
    time.sleep(3)

    # Find the email and password fields and enter the credentials
    email_field = driver.find_element(By.NAME, "email")
    email_field.send_keys("joshycash20023@gmail.com")

    password_field = driver.find_element(By.NAME, "password")
    password_field.send_keys("Samwassef123!")
    password_field.send_keys(Keys.RETURN)  # Press Enter to submit
    print("Logged in successfully.")

    # Wait for the login process to complete
    time.sleep(5)

    # Navigate to the property page (initially lands on the /building tab)
    renonomy_link = "https://app.reonomy.com/!/property/d91a0879-2645-5aa4-874d-d94152d7327a/building"
    driver.get(renonomy_link)
    print("Navigated to the property page.")

    # Wait for the page to load completely
    time.sleep(5)

    # Click on the "Owner" tab
    owner_tab = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "Owner"))
    )
    owner_tab.click()
    time.sleep(3)  # Wait for the owner tab content to load
    print("Navigated to the owner tab.")

    # Extract and organize the data
    try:
        # Locate the start element
        start_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//p[@data-testid='header-property-address']"))
        )
        print("Start element found: ", start_element.text)

        # Scroll to the start element
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", start_element)
        time.sleep(1)

        # Initialize organized data lists
        address = start_element.text
        apn = ""
        names = []
        phone_numbers = []
        extra_addresses = []
        emails = []

        # Locate all paragraph elements following the start element
        elements_in_between = start_element.find_elements(By.XPATH, "following::p")
        for element in elements_in_between:
            text = element.text
            if "APN:" in text:
                apn = text
            elif any(keyword in text.lower() for keyword in ["owner", "reported owner"]):
                names.append(text)
            elif "phone" in text.lower() or text.isdigit():
                phone_numbers.append(text)
            elif "address" in text.lower() or "st" in text.lower():
                extra_addresses.append(text)
            elif "@" in text:
                emails.append(text)

        # Organize the extracted data
        organized_text = f"{address}\n{apn}\n\nNames:\n" + "\n".join(names) + \
                         "\n\nPhone Numbers:\n" + "\n".join([f"• {number}" for number in phone_numbers]) + \
                         "\n\nExtra Addresses:\n" + "\n".join([f"• {addr}" for addr in extra_addresses]) + \
                         "\n\nEmails:\n" + "\n".join([f"• {email}" for email in emails])

        print("Text organized and extracted successfully.")

        # Path to the service account JSON key file
        SERVICE_ACCOUNT_FILE = '/Users/joshpirrie/Downloads/reon-export-3aa15e25d698.json'

        # Define the scopes required
        SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

        # Authenticate and build the Google Docs service
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('docs', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        # Create a new Google Doc
        document = service.documents().create(body={
            'title': 'Extracted Owner Information'
        }).execute()

        # Get the document ID
        document_id = document['documentId']

        # Prepare the text to be inserted
        requests = [
            {
                'insertText': {
                    'location': {
                        'index': 1,
                    },
                    'text': organized_text
                }
            }
        ]

        # Use the batchUpdate method to insert the text into the document
        result = service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

        print(f"Text successfully saved to Google Doc: {document['title']}")

        # Share the document with your personal Google account
        user_permission = {
            'type': 'user',
            'role': 'writer',
            'emailAddress': 'jpirrie63@gmail.com'
        }
        drive_service.permissions().create(
            fileId=document_id,
            body=user_permission,
            fields='id',
        ).execute()

        # Open the Google Doc in the default web browser
        doc_url = f"https://docs.google.com/document/d/{document_id}/edit"
        webbrowser.open(doc_url)
        print("Google Doc opened in browser.")

        # Pause to ensure the document is loaded in the browser before making the final updates
        time.sleep(5)

        # Insert "DATA EXTRACTED, READY FOR NEXT PROPERTY" into the document with green highlight
        final_requests = [
            {
                'insertText': {
                    'location': {
                        'index': len(organized_text) + 1,
                    },
                    'text': "\n\nDATA EXTRACTED, READY FOR NEXT PROPERTY"
                }
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': len(organized_text) + 3,
                        'endIndex': len(organized_text) + 44
                    },
                    'textStyle': {
                        'backgroundColor': {
                            'color': {
                                'rgbColor': {
                                    'red': 0.0,
                                    'green': 1.0,
                                    'blue': 0.0
                                }
                            }
                        }
                    },
                    'fields': 'backgroundColor'
                }
            }
        ]

        # Update the document with the final text
        service.documents().batchUpdate(documentId=document_id, body={'requests': final_requests}).execute()
        print("Final text added to the document.")

    except Exception as e:
        print("An error occurred while extracting owner information:", e)

except Exception as e:
    print("An error occurred:", e)

finally:
    # Close the driver
    driver.quit()
    print("Script completed.")
