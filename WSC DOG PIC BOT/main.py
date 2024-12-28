import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Set up the Chrome driver options
options = Options()
options.add_argument("profile-directory=Profile 1")
options.add_argument(r"C:\Users\THAAGAM\Desktop\WSC\chrome-data")  # Use raw string to avoid unicode error
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# Start the browser with the given options
driver = webdriver.Chrome(options=options)
driver.get("http://web.whatsapp.com")

# Wait for the user to scan the QR code
input("Press Enter after scanning QR code")

def send_whatsapp_msg(driver, phone_no, text):
    """Function to send a text message to a specific phone number."""
    try:
        # Wait for the message input field to be visible
        txt_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p'))
        )
        txt_box.send_keys(text)
        txt_box.send_keys(Keys.ENTER)
        print(f"Message sent: {text}")
    except Exception as e:
        print(f"Error sending message: {e}")

def get_last_message(driver):
    """Function to get the last message from the chat."""
    try:
        # Wait until the message container is loaded
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[3]/div/div[2]/div[3]/div'))
        )

        # Get all message elements in the chat container
        message_elements = driver.find_elements(By.XPATH, '//*[@id="main"]/div[3]/div/div[2]/div[3]/div/div')

        # Get the last message from the list
        if message_elements:
            last_message = message_elements[-1].text  # Last message in the list
            print(f"Last message: {last_message}")
            return last_message
        else:
            print("No messages found.")
            return None
    except Exception as e:
        print(f"Error while retrieving the last message: {e}")
        return None

def get_phone_number(driver):
    """Function to get the phone number or contact name of the currently opened chat."""
    try:
        # Click on the header to reveal the contact information
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/header/div[2]/div[1]/div'))
        ).click()

        # Extract the phone number or contact name from the header
        contact_info = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/header/div[2]/div[2]/span'))
        ).text

        print(f"Phone number or contact name: {contact_info}")
        return contact_info
    except Exception as e:
        print(f"Error while retrieving the contact info: {e}")
        return None

def open_unread_chat(driver):
    """Function to open unread chats by clicking the unread button."""
    try:
        # Click on the unread button (//*[@id="side"]/div[2]/button[2])
        unread_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="side"]/div[2]/button[2]'))
        )
        unread_button.click()
        time.sleep(2)  # Wait for the sidebar to open
        
        # Click the first unread chat (//*[@id="pane-side"]/div[1]/div/div/div[1]/div/div/div/div[2])
        unread_chat = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="pane-side"]/div[1]/div/div/div[1]/div/div/div/div[2]'))
        )
        unread_chat.click()
        return True
    except Exception as e:
        print(f"Error while opening unread chat: {e}")
        return False

# Start the bot loop
while True:
    if open_unread_chat(driver):  # Open the first unread chat
        # Wait for the chat to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[3]/div/div/div[2]/div[1]/div[1]'))
        )

        # Retrieve the last message in the chat
        last_message = get_last_message(driver)

        if last_message:
            # Add logic for automatic replies based on the message content
            if "hi" in last_message.lower():
                send_whatsapp_msg(driver, get_phone_number(driver), "Hello Sir/Madam.... I am Boobalan (+91782392353) Your Donor Relationship Manager. Would you like to donate?")
                print("Replied with donation message.")
            elif "yes" in last_message.lower() or "yeah" in last_message.lower():
                send_whatsapp_msg(driver, get_phone_number(driver), "Ok, is there any special occasion?")
                print("Replied with occasion question.")
            elif "birthday" in last_message.lower() or "anniversary" in last_message.lower():
                send_whatsapp_msg(driver, get_phone_number(driver), "Here is the plan!")
                print("Replied with plan message.")
            elif "no" in last_message.lower():
                send_whatsapp_msg(driver, get_phone_number(driver), "Okay, thanks for your time!")
                print("Replied with thanks message.")
    
    # Sleep for a few seconds before checking again
    time.sleep(5)
