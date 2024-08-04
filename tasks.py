from robocorp.tasks import task
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from googletrans import Translator
from urllib.parse import urlparse
import re
import pandas as pd
import time
import os
import requests
import subprocess
import json
import logging
import sys


"""For the following project we will access a G1 news website, and within it we will carry out a search,
and after the search it will perform filters on the page such as: category and period, all data removed by work items 
Following the procedure, it will extract the information from each news item presented: title, description (if any), 
date (if any), image (if any), and after extraction and download it will perform analyzes such as: count of phrases 
search on title and description, True or false depending on whether the title or description contains any amount 
of money (USD and BRL), after completing all the news, it will generate an excel file with all the information"""

"""To change the default search mode, change devdata/work-items-in/work_items.json/work-items.json
    keyword = ''
    category = ['All results', 'news', 'photos', 'Videos', 'blogs']
    filter_date= ['Anytime', 'At the last minute', 'In the last 24 hours', 'Last week', 'In the last month','In the last year']"""




# set variables
folder = "output/"
spreadsheet = "News_Reports.xlsx"


# Main function
@task
def main():
    # validates browser and webdriver
    run_script("script_browser.py")
    # Browser configuration
    log_path = "logs/logs.txt"
    exec_path = 'C:/ProgramData/msedgedriver.exe'
    services = webdriver.EdgeService(log_output=log_path, executable_path=exec_path)
    driver = webdriver.Edge(service=services)

    # import work items
    with open("devdata/work-items-in/work_items.json/work-items.json", "r") as file:
        current_work_item = json.load(file)

    item = current_work_item["payload"]

    for i in item:
        keyword = i["keyword"] if i["keyword"] else "money"
        category = i["category"] if i["category"] else "news"
        category_translated = translated(category)
        filter_date = i["filter_date"] if i["filter_date"] else "In the last 24 hours"
        filter_date_translated = translated(filter_date)

    # validate output directory
    if not os.path.exists(folder):
        os.makedirs(folder)

    # start process
    driver.get("https://g1.globo.com/")

    # maxime browser
    driver.maximize_window()

    # wait if there is a cookie button
    try:
        if WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable(
                (By.CLASS_NAME, "cookie-banner-lgpd_accept-button")
            )
        ):
            driver.find_element(
                By.CLASS_NAME, "cookie-banner-lgpd_accept-button"
            ).click()
    except Exception as e:
        print(f"Error clicking accept cookies")

    # Wait until the search field is available and perform the search
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="search"]'))
    )
    search_box = driver.find_element(By.CSS_SELECTOR, 'input[type="search"]')
    search_box.send_keys(f"{keyword}")
    search_box.send_keys(Keys.RETURN)

    # Wait for the results to load
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CLASS_NAME, "results__list"))
    )

    # Select the web item category
    driver.find_element(
        By.XPATH,
        '//*[@id="search-filter-component"]/div/div[1]/div/div/div[1]/div[1]/a/span[2]',
    ).click()
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                f'//span[contains(text(),"{category}") or contains(text(),"{category_translated}")]',
            )
        )
    )
    driver.find_element(
        By.XPATH,
        f'//span[contains(text(),"{category}") or contains(text(),"{category_translated}")]',
    ).click()

    # Select the order by
    driver.find_element(
        By.XPATH,
        '//*[@id="search-filter-component"]/div/div[1]/div/div/div[1]/div[2]/a/span[2]',
    ).click()

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Recente")]'))
    )
    driver.find_element(
        By.XPATH, '//span[contains(text(),"Recente") or contains(text(),"Recent")]'
    ).click()

    # Filter by work item or 'Last Week'
    driver.find_element(
        By.XPATH,
        '//*[@id="search-filter-component"]/div/div[1]/div/div/div[2]/div/a/span[2]',
    ).click()
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                f'//span[contains(text(),"{filter_date}") or contains(text(),"{filter_date_translated}")]',
            )
        )
    )
    driver.find_element(
        By.XPATH,
        '//span[contains(text(),"Na última semana") or contains(text(),"Last Week")]',
    ).click()

    # Wait for the news to load
    time.sleep(7)  # Wait 7 seconds to ensure the page has fully loaded

    # Scroll to the end of the page to load all news
    scroll_to_end(driver)

    # Extract news information
    results = []

    news_items = driver.find_elements(
        By.XPATH, '//li[contains(@class, "widget--card")]'
    )

    for item in news_items:

        WebDriverWait(item, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "a"))
        )

        # Image URL (if available)
        try:
            images = item.find_elements(By.TAG_NAME, "img")
            image_filename = ""
            for image_download in images:
                try:
                    image_url = image_download.get_attribute("src")
                    image_download = download_image(image_url, folder)
                    if (
                        image_filename
                    ):  # If the variable already has a value, add a semicolon
                        image_filename += ";"
                        image_filename += image_download
                    else:
                        image_filename = image_download
                except:
                    None
        except Exception as e:
            print(f"Error processing item: {e}")
            image_filename = None

        # Open the news in a new tab
        link_element = item.find_element(By.TAG_NAME, "a")
        news_url = link_element.get_attribute("href")

        driver.execute_script("window.open(arguments[0], '_blank');", news_url)
        driver.switch_to.window(driver.window_handles[-1])

        try:
            # Publication date
            date_published = driver.find_element(
                By.XPATH, '//time[contains(@itemprop, "datePublished")]'
            ).text
            date = parse_date(date_published)

            # Full title
            title_element = driver.find_element(
                By.XPATH, '//h1[contains(@class, "content-head__title")]'
            )
            title = title_element.text
            title_count = count_occurrences(title, keyword)
            title_contains_money = contains_money(title)

            # Full description (if available)
            try:
                description_element = driver.find_element(
                    By.XPATH, '//h2[contains(@class, "content-head__subtitle")]'
                )
                description = description_element.text
            except:
                description = None

            description_count = count_occurrences(description, keyword)
            description_contains_money = contains_money(description)

            # Add the data to the results list
            results.append(
                {
                    "date": date,
                    "title": title,
                    "title_count": title_count,
                    "title_contains_money": title_contains_money,
                    "description": description,
                    "description_count": description_count,
                    "description_contains_money": description_contains_money,
                    "image_filename": image_filename,
                }
            )

        except Exception as e:
            print(f"Error processing item: {e}")

        # Close the tab and return to the main tab
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    # Save the results to an Excel spreadsheet
    try:
        df = pd.DataFrame(results)
        df.to_excel(folder + spreadsheet, index=False)
    except Exception as e:
        print(f"Error processing: {e}")

    # Close the browser
    driver.quit()


# function to validate the browser
def run_script(script_name):
    try:
        result = subprocess.run(
            ["python", script_name], check=True, text=True, capture_output=True
        )
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"Error executing {script_name}: {e}")
        print(e.output)


# function to translate word
def translated(text, language="pt"):
    translator = Translator()
    try:
        translation = translator.translate(text, dest=language)
        return translation.text
    except Exception as e:
        print(f"Wrong during translation: {e}")
        return text  # Return the originaal word


# function to convert date
def parse_date(date):

    try:
        date = datetime.strptime(date.replace("h", ":"), "%d/%m/%Y %H:%M").strftime(
            "%d/%m/%Y"
        )
    except Exception as e:
        ValueError("The string does not match the expected format.")
        date = None
    return date


# Function to download an image and save it locally
def download_image(url, folders):

    folder_path = folders + "news_photos/"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    try:
        # Request image
        response = requests.get(url, stream=True)
        response.raise_for_status()

        # Check if is a image
        if "image" in response.headers.get("Content-Type", ""):
            # Get the image name
            filename = os.path.basename(urlparse(url).path)
            file_path = os.path.join(folder_path, filename)

            # Save on directory
            with open(file_path, "wb") as file:
                file.write(response.content)

            return filename
        else:
            print("The URL provided is not an image.")
            return None

    except requests.RequestException as e:
        print(f"Error downloading image: {e}")
        return None


# Function to count keyword occurrences
def count_occurrences(text, keyword):
    return text.lower().count(keyword.lower())


# Function to check if the text contains a money amount (USD | BRL)
def contains_money(text):
    money_pattern = re.compile(
        r"\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+\s?(?:dollars|USD)|R\$\s?\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?|\d+\s?(?:reais|BRL)",
        re.IGNORECASE,
    )
    return bool(money_pattern.search(text))


# Function to scroll to the end of the page
def scroll_to_end(driver):
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)  # Wait to load new news
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


if __name__ == "__main__":
    main()
