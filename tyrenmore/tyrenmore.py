import time
import random
import re
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, load_workbook
from webdriver_manager.chrome import ChromeDriverManager

def get_tyre_price_amazon(model, width, profile, rim):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:

        search_query = f"{model} {width}/{profile}R{rim} tyre"
        url = f"https://www.amazon.in/s?k={search_query.replace(' ', '+')}"
        print(f"Fetching data from Amazon: {url}")

        # Open the URL
        driver.get(url)

        # Wait for the search results to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.s-main-slot"))
        )

        # Extract price and link for the first matching result
        items = driver.find_elements(By.CSS_SELECTOR, "div.s-main-slot div[data-component-type='s-search-result']")
        for item in items:
            try:
                price_tag = item.find_element(By.CSS_SELECTOR, "span.a-price-whole")
                link_tag = item.find_element(By.CSS_SELECTOR, "a.a-link-normal")

                if price_tag and link_tag:
                    price = int(price_tag.text.replace(",", "").strip())
                    link = link_tag.get_attribute("href")  # Corrected line to avoid duplication
                    return price, link
            except Exception as e:
                # Skip any items that don't have the required elements
                print("Skipping an item due to error:", e)
                continue

        return None, None

    finally:
        driver.quit()



def get_tyre_price_flipkart(model, width, profile, rim):
    # Set up Selenium WebDriver with Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Use webdriver-manager to download and manage ChromeDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        # Build the search query
        search_query = f"{model} {width}/{profile}R{rim} tyre"
        url = f"https://www.flipkart.com/search?q={search_query.replace(' ', '+')}"
        print(f"Fetching data from Flipkart: {url}")

        # Open the URL
        driver.get(url)

        # Wait for the page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "slAVV4"))  # Wait for product containers
        )

        # Scroll to ensure all products are loaded
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Locate the product containers
        items = driver.find_elements(By.CLASS_NAME, "slAVV4")  # Class for product containers

        for item in items:
            try:
                # Find the price and link within the product container
                price_tag = item.find_element(By.CLASS_NAME, "Nx9bqj")  # Class for price
                link_tag = item.find_element(By.TAG_NAME, "a")  # Tag for link

                if price_tag and link_tag:
                    # Extract price and link
                    price = int(price_tag.text.replace("₹", "").replace(",", "").strip())
                    link = link_tag.get_attribute("href")
                    if not link.startswith("http"):
                        link = "https://www.flipkart.com" + link
                    return price, link
            except Exception as e:
                # Skip items without price or link
                print("Skipping an item due to error:", e)
                continue

        # Return None if no products were found
        return None, None

    except Exception as e:
        print(f"Error fetching data from Flipkart: {e}")
        return None, None

    finally:
        driver.quit()


# Function to read links from an Excel file
def read_links_from_excel(file_name="input_link.xlsx"):
    try:
        df = pd.read_excel(file_name)  # Read Excel file
        return df['Links'].dropna().tolist()  # Extract the column named 'Links' and convert to list
    except Exception as e:
        print(f"Error reading links from Excel: {e}")
        return []
    




def scroll_and_load_all_products(driver):
    """ Scrolls and clicks 'Load More' button until all products are fully loaded. """
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        # ✅ Scroll to the bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)  # Give time for content to load

        # ✅ Try clicking "Load More" button
        try:
            load_more_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "load-more"))
            )

            if load_more_button.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView();", load_more_button)  # Ensure it's in view
                time.sleep(1)  # Allow time to adjust
                driver.execute_script("arguments[0].click();", load_more_button)  # Click using JavaScript
                time.sleep(5)  # Wait for new content to load

                # ✅ Scroll again after clicking "Load More" to trigger loading
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(3)  # Give time for additional content to load

        except Exception as e:
            print(f"Load More button not found or not clickable: {e}")
            pass  # No 'Load More' button found, continue scrolling

        # ✅ Check if new content has loaded
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:  # If height doesn’t change, stop
            break  
        last_height = new_height

def scrape_tyres(url):
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")  # Modern headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--log-level=3")  # Reduce logs
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Anti-detection
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    # Configure Selenium WebDriver
    chrome_options = Options()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    chrome_options.add_argument("--headless") 

    try:
        # Open the website
        driver.get(url)

        # Wait for the page to load and for tyre models to appear
        WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "product-item-link")))
        scroll_and_load_all_products(driver)

        # Extract tyre model names and links from the page
        tyres = []  # List to store tyre details as dictionaries

        try:
            # Fetch all model elements
            model_elements = driver.find_elements(By.CLASS_NAME, 'product-item-link')

            for model_element in model_elements:
                model_name = model_element.text.strip()
                link = model_element.get_attribute("href")  # Fetch product link
                if model_name and link:  # Ensure both model name and link exist
                    tyres.append({"model_name": model_name, "link": link, "price": None, "features": None, "warranty_part1": None, "warranty_part2": None})
        except Exception as e:
            print(f"Error extracting tyre models or links: {e}")

        # Extract prices, features, and warranty by visiting individual product pages
        for tyre in tyres:
            try:
                # Navigate to the product page
                driver.get(tyre["link"])

                # Extract price
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "price-wrapper")))
                    price_element = driver.find_element(By.CLASS_NAME, "price-wrapper")
                    tyre["price"] = float(price_element.get_attribute("data-price-amount"))  # Add price to the dictionary
                except Exception as e:
                    print(f"Error fetching price for model '{tyre['model_name']}': {e}")

                # Extract features
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "item")))
                    feature_element = driver.find_element(By.ID, "Features")
                    tyre["features"] = feature_element.get_attribute("innerText").strip()  # Add features to the dictionary
                except Exception as e:
                    print(f"Error fetching features for model '{tyre['model_name']}': {e}")

                # Extract warranty information
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Warranty")))

                    # Find all elements with the class name "warrnty-info"
                    warranty_elements = driver.find_elements(By.CLASS_NAME, "warrnty-info")

                    # Extract text from each WebElement
                    warranties = [element.text.strip() for element in warranty_elements]

                    # Store the first warranty in "warranty_part1" and the second in "warranty_part2"
                    tyre["warranty_part1"] = warranties[0] if len(warranties) > 0 else "N/A"
                    tyre["warranty_part2"] = warranties[1] if len(warranties) > 1 else "N/A"

                except Exception as e:
                    print(f"Error fetching warranty for model '{tyre['model_name']}': {e}")
                
                # # Get Amazon price and link for the tyre
                model = tyre["model_name"]
                width, profile, rim = None, None, None  # You might extract this data if needed

                amazon_price, amazon_link = get_tyre_price_amazon(model, width, profile, rim)
                tyre["amazon_price"] = amazon_price
                tyre["amazon_link"] = amazon_link


                flipkart_price, flipkart_link = get_tyre_price_flipkart(model, width, profile, rim)
                tyre["flipkart_price"] = flipkart_price
                tyre["flipkart_link"] = flipkart_link


            except Exception as e:
                print(f"Error processing tyre '{tyre['model_name']}': {e}")

        # Store tyre models, prices, links, features, and warranties in the Excel file
        save_models_to_excel(tyres)

    except Exception as e:
        print(f"An error occurred while scraping brand page: {e}")

    finally:
        driver.quit()

def save_models_to_excel(tyres):
    file_name = "product.xlsx"

    try:
        workbook = load_workbook(file_name)
        sheet = workbook.active

        sheet["B1"] = "Models"
        sheet["C1"] = "Patterns"
        sheet["D1"] = "Width"
        sheet["E1"] = "Profile"
        sheet["F1"] = "Contaminant Radius"
        sheet["G1"] = "Load Rating"
        sheet["H1"] = "Speed Rating"
        sheet["I1"] = "Type"
        sheet["S1"] = "FEATURE 1"  # Updated column for A/T or AT
        sheet["T1"] = "FEATURE 2"  # Updated column for L/T or LT
        sheet["U1"] = "FEATURE 3" 
        sheet["AJ1"] = "Tyrenmore Price"
        sheet["AK1"] = "Tyrenmore Link"
        sheet["AL1"] = "Amazon Price"
        sheet["AM1"] = "Amazon Link"
        sheet["AN1"] = "Price of Flipkart"
        sheet["AO1"] = "Url of Flipkart"   
        sheet["V1"] = "Ride"
        sheet["W1"] = "Braking"  # Updated to store braking feature
        sheet["X1"] = "Grip"     # Updated to store grip feature
        sheet["Y1"] = "Stability" # Updated to store stability feature
        sheet["Z1"] = "noise",
        sheet["AA1"] = "fuel",
        sheet["AD1"] = "Conditional"
        sheet["AE1"] = "Unconditional"

        # Write each model name, price, link, features, and warranty parts into respective columns
        for row, tyre in enumerate(tyres, start=2):
            model = tyre.get("model_name")
            price = tyre.get("price")
            link = tyre.get("link")
            amazon_price = tyre.get("amazon_price")
            amazon_link = tyre.get("amazon_link")
            flipkart_price = tyre.get("flipkart_price")
            flipkart_link = tyre.get("flipkart_link")
            features = tyre.get("features")
            warranty_part1 = tyre.get("warranty_part1")
            warranty_part2 = tyre.get("warranty_part2")

            # Initialize values for columns S, T, and U
            at_value = "N/A"
            lt_value = "N/A"
            uhp_value = "N/A"

            # Split the model into parts and check for terms
            model_parts = model.upper().split() if model else []

            if "A/T" in model_parts or "AT" in model_parts:
                at_value = "AT"
            if "L/T" in model_parts or "LT" in model_parts:
                lt_value = "L/T or LT"
            if "UHP" in model_parts:
                uhp_value = "UHP"

            # Save the extracted values into the respective columns
            sheet[f"S{row}"] = at_value
            sheet[f"T{row}"] = lt_value
            sheet[f"U{row}"] = uhp_value

            # Split features into Braking, Grip, and Stability
            ride = "N/A"
            braking = "N/A"
            grip = "N/A"
            stability = "N/A"
            noise = "N/A"
            fuel = "N/A"

            if features:
                if "ride" in features.lower():
                    ride  = "Smooth ride"
                if "braking" in features.lower():
                    braking = "Excellent braking"
                if "grip" in features.lower():
                    grip = "Excellent Dry & Wet Grip"
                if "stable" in features.lower():
                    stability = "Highly stable"
                if "noise" in features.lower():
                    noise="Low Noise"
                if "fuel" in features.lower():
                    fuel="Fuel Eficiency"

            # Save features into respective columns
            sheet[f"V{row}"] = ride
            sheet[f"W{row}"] = braking
            sheet[f"X{row}"] = grip
            sheet[f"Y{row}"] = stability
            sheet[f"Z{row}"] = noise
            sheet[f"AA{row}"] = fuel

            # Save other details
            sheet[f"B{row}"] = model
            sheet[f"AJ{row}"] = price if price is not None else "N/A"
            sheet[f"AK{row}"] = link
            sheet[f"AL{row}"] = amazon_price if amazon_price is not None else "N/A"
            sheet[f"AM{row}"] = amazon_link if amazon_link is not None else "N/A"
            sheet[f"AN{row}"] = flipkart_price if flipkart_price is not None else "N/A"
            sheet[f"AO{row}"] = flipkart_link if flipkart_link is not None else "N/A"
            sheet[f"AD{row}"] = warranty_part1 if warranty_part1 else "N/A"
            sheet[f"AE{row}"] = warranty_part2 if warranty_part2 else "N/A"

            # Process and extract tyre details
            tyre_type = None 
            width, profile, rim, load_rating, speed_rating = None, None, None, None, None

            # Split the model into parts
            if model:
                model_parts = model.split()

            # Identify tyre type (Tubeless or Tube-type)
            if "TUBELESS" in model.upper() if model else False:
                tyre_type = "Tubeless"
            elif "TUBE-TYPE" in model.upper() if model else False:
                tyre_type = "Tube-type"

            # Look for tyre size patterns (e.g., "145/80 R12", "100/90-19", "110/80-14")
            for part in model_parts:
                if "/" in part and "-" in part:  # Handle cases like "110/80-14"
                    try:
                        width_profile, rim = part.split("-")  # Split at the hyphen
                        width, profile = width_profile.split("/")  # Split width and profile
                        if width.isdigit() and profile.isdigit() and rim.isdigit():
                            sheet[f"D{row}"] = int(width)
                            sheet[f"E{row}"] = int(profile)
                            sheet[f"F{row}"] = int(rim)
                    except ValueError:
                        pass
                elif "/" in part:  # Handle cases like "145/80 R12"
                    try:
                        width, profile = part.split("/")
                        if width.isdigit() and profile.isdigit():
                            sheet[f"D{row}"] = int(width)
                            sheet[f"E{row}"] = int(profile)
                    except ValueError:
                        pass
                elif part.startswith("R") and part[1:].isdigit():  # Handle cases like "R12"
                    sheet[f"F{row}"] = int(part[1:])

            # Extract load and speed ratings (e.g., "82T")
            for part in model_parts:
                if part[-1].isalpha() and part[:-1].isdigit():
                    load_rating = int(part[:-1])
                    speed_rating = part[-1]
                    break
                match = re.search(r"(6i).*?(\d+)\s?([A-Z])", part, re.IGNORECASE)
                if match:
                    load_rating = int(match.group(2))  # Extract load index as integer
                    speed_rating = match.group(3)     # Extract speed rating
                    break

                
            # Save extracted details
            sheet[f"I{row}"] = tyre_type
            sheet[f"G{row}"] = load_rating
            sheet[f"H{row}"] = speed_rating

        # Save the workbook
        workbook.save(file_name)
        print(f"Data successfully saved to {file_name}")

    except Exception as e:
        print(f"Error saving model names to Excel: {e}")



if __name__ == "__main__":
    # url=input("Enter the Website url: ")
    # scrape_tyres(url)
    input_file = "input_links.xlsx"
    urls = read_links_from_excel(input_file)
    count=1

    if not urls:
        print("No valid links found in the input file.")
    else:
        print("Started")

        for url in urls:
            print(f"{count}. Scraping: {url}")
            count=count+1
            time.sleep(random.randint(3,6))  # Random delay to avoid blocking
            tyres = scrape_tyres(url)
