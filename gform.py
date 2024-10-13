import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# URL of the Google Form
form_link = "https://forms.gle/gFX7TmZBBcrW69Up8"

# Sample data to be submitted (you can extract this from a CSV or an API)
data_to_submit = [
    {
        "Customer Name": "John Doe",
        "Order ID": "132",
        "Feedback Category": "Return/Exchange - Damaged",
        "Sentiment": "Negative"  # This can be "Positive", "Neutral", or "Negative"
    },
    {
        "Customer Name": "Jane Smith",
        "Order ID": "297",
        "Feedback Category": "Product Issue",
        "Sentiment": "Positive"
    },
    # Add more entries as needed
]

def main():
    # Configure Chrome options
    options = webdriver.ChromeOptions()
    options.add_argument("-incognito")
    # Uncomment the line below for headless mode
    # options.add_argument("--headless")

    # Create the Chrome WebDriver
    browser = webdriver.Chrome(options=options)

    try:
        # Open the Google Form
        browser.get(form_link)

        for entry in data_to_submit:
            print(f"Entering data for: {entry['Customer Name']}")
            
            # Wait and locate the "Customer Name" input field
            customer_name_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//input[@type='text' and @aria-labelledby='i1']"))
            )
            customer_name_field.clear()
            customer_name_field.send_keys(entry["Customer Name"])
            print("Entered Customer Name")

            # Wait and locate the "Order ID" input field
            order_id_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//input[@type='text' and @aria-labelledby='i5']"))
            )
            order_id_field.clear()
            order_id_field.send_keys(entry["Order ID"])
            print("Entered Order ID")

            # Wait and locate the "Feedback Category" input field
            feedback_category_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//input[@type='text' and @aria-labelledby='i9']"))
            )
            feedback_category_field.clear()
            feedback_category_field.send_keys(entry["Feedback Category"])
            print("Entered Feedback Category")

            # Handling Sentiment as radio buttons
            sentiment = entry["Sentiment"]
            if sentiment == "Positive":
                sentiment_field = WebDriverWait(browser, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//div[@class='nWQGrd zwllIb']//*[text()='Positive']"))
                )
            elif sentiment == "Neutral":
                sentiment_field = WebDriverWait(browser, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//div[@class='nWQGrd zwllIb']//*[text()='Neutral']"))
                )
            elif sentiment == "Negative":
                sentiment_field = WebDriverWait(browser, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//div[@class='nWQGrd zwllIb']//*[text()='Negative']"))
                )
            sentiment_field.click()
            print(f"Selected Sentiment: {sentiment}")

            # Locate and click the submit button
            submit_button = WebDriverWait(browser, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']"))
            )
            submit_button.click()
            print("Form submitted successfully")

            # Wait for the form to reload before the next submission
            time.sleep(3)  # Adjust timing as necessary to avoid bot detection

            # Re-open the form for the next entry
            browser.get(form_link)

    finally:
        # Close the browser
        browser.quit()

if __name__ == "__main__":
    main()
