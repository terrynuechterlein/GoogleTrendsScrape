from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time

# Setup Google Sheets access
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("C:/Users/terry/Programming/Python/json_key/conductive-fold-413909-fe0dab1a6107.json", scope)
client = gspread.authorize(creds)
sheet = client.open('GoogleTrends').sheet1

# Setup Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run Chrome in headless mode
service = ChromeService(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Navigate to Google Trends and wait for the page to load
driver.get("https://trends.google.com/trends/trendingsearches/daily?geo=US")
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@class="details-top"]//a')))

# Find the trending searches elements
trends = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@class="details-top"]//a')))
titles = [trend.text for trend in trends[:20]]

# Update Google Sheet
for i, (title, search_count) in enumerate(zip(titles), start=1):
    sheet.update_cell(i + 1, 1, str(i))  # Update row number in column 1
    sheet.update_cell(i + 1, 2, title)  # Update title in column 2
    sheet.update_cell(i + 1, 3, search_count)  # Update search count in column 3

driver.quit()