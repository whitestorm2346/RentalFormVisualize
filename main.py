import pandas as pd
from selenium import webdriver  # for operating the website
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import chromedriver_autoinstaller
import folium
import re
from time import sleep

DATA_FILE = "name.xlsx"
GOOGLE_MAP = "https://www.google.com.tw/maps"

def extract_lat_lng(url):
    # print(url)
    match = re.search(r'@(-?\d+\.\d+),(-?\d+\.\d+)', url)

    if match:
        lat, lng = match.groups()
        return (float(lat), float(lng))
    
    return None

def get_lat_lng(driver, address):
    driver.get(GOOGLE_MAP)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    driver.implicitly_wait(3)

    search_box_input = driver.find_element(By.XPATH, '//*[@id="searchboxinput"]')
    search_box_input.send_keys(address)

    search_button = driver.find_element(By.XPATH, '//*[@id="searchbox-searchbutton"]')
    search_button.click()

    while driver.current_url == GOOGLE_MAP:
        pass

    new_url = driver.current_url

    driver.refresh() # 確保大頭針在視窗中央 -> 確保經緯度正確

    while driver.current_url == new_url:
        pass

    return extract_lat_lng(driver.current_url)
    

if __name__ == "__main__":
    try:
        df = pd.read_excel(DATA_FILE)
        addresses = df['地址'].tolist()

        chrome_option = chromeOptions()
        chrome_option.add_argument('--log-level=3')
        chrome_option.add_argument('--start-maximized')
        chromedriver_autoinstaller.install()
        driver = webdriver.Chrome(options=chrome_option)

        coordinates = []

        for address in addresses:
            coord = get_lat_lng(driver, address)
            # print(coord)
            coordinates.append(coord)
            sleep(1)  

        tku_coord = [25.174542, 121.450259] 
        mymap = folium.Map(location=tku_coord, zoom_start=16, tiles="OpenStreetMap")

        for address, coord in zip(addresses, coordinates):
            if coord:
                folium.Marker(location=coord, popup=address).add_to(mymap)

        mymap.save('map.html')
        print("地圖已保存為 'map.html'")
    except Exception as e:
        print('Excel file not found!')