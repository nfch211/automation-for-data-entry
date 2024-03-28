from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import random

def extract_page_content(url):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    driver.get(url)

    # Scroll down
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(random.randint(2,5))

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        soup = BeautifulSoup(driver.page_source, "html.parser")
        content = soup.get_text()

        return content
    except Exception as e:
        print(f"Failed to retrieve content from: {url}. Error: {str(e)}")
        return None
    finally:
        driver.quit()

def search_br_number_on_baidu(br_number):
    url = f"https://www.baidu.com/s?wd={br_number}"

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    driver.get(url)
    time.sleep(random.randint(2,5))  # wait before searching

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".c-container h3 a"))
        )

        result_items = driver.find_elements(By.CSS_SELECTOR, ".c-container h3 a")
        if result_items:
            print(f"Search results for BR Number {br_number} on Baidu:\n")
            for idx, item in enumerate(result_items, 1):
                title = item.text
                link = item.get_attribute("href")
                print(f"{idx}. {title}\n   {link}")

                time.sleep(random.randint(2,5))  # wait before accessing link

                content = extract_page_content(link)
                if content:
                    print(f"\nContent of the web page:\n{content}\n")
                else:
                    print("Failed to retrieve content.")
        else:
            print("No results found.")
    except Exception as e:
        print(f"Failed to retrieve search results. Error: {str(e)}")
    finally:
        driver.quit()

if __name__ == "__main__":
    br_number_to_search = "91440101MA59KR6723"
    search_br_number_on_baidu(br_number_to_search)
