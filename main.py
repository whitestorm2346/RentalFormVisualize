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

# DATA_FILE = "113學年度學生居住調查暨校外賃居安全自評表_部份資料測試用.xlsx"
DATA_FILE = "113學年度學生居住調查暨校外賃居安全自評表.xlsx"
GOOGLE_MAP = "https://www.google.com.tw/maps"


ID_COL_IDX = 64
FINISH_DT_COL_IDX = 2
CURRENT_RESIDENSY_COL_IDX = 7
ADDRESS_COL_IDX = 9
PROPERTY_LABEL_COL_IDX = 10
PROPERTY_TYPE_COL_IDX = 14


def remove_duplicate(title, data):
    # data[title[FINISH_DT_COL_IDX]] = pd.to_datetime(data[title[FINISH_DT_COL_IDX]], format='%m/%d/%y %H:%M')
    data[title[FINISH_DT_COL_IDX]] = pd.to_datetime(data[title[FINISH_DT_COL_IDX]], format='%m/%d/%y %H:%M:%S')
    latest_entries = data.loc[data.groupby(title[ID_COL_IDX])[title[FINISH_DT_COL_IDX]].idxmax()]

    return latest_entries

def residency_filter(title, data):
    filtered_data = data[data[title[CURRENT_RESIDENSY_COL_IDX]].str.contains('校外租屋', na=False)]

    return filtered_data

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

    sleep(1)
    new_url = driver.current_url

    driver.refresh() # 確保大頭針在視窗中央 -> 確保經緯度正確

    while driver.current_url == new_url:
        pass

    return extract_lat_lng(driver.current_url)
    

if __name__ == "__main__":
    try:
        df = pd.read_excel(DATA_FILE, header=None)
    except Exception as e:
        print('Excel file not found!')
        exit(1)

    title = df.iloc[0, :]
    data = df.iloc[1:, :]
    data.columns = title

    latest_entries = remove_duplicate(title, data) 
    filtered_data = residency_filter(title, latest_entries)

    chrome_option = chromeOptions()
    chrome_option.add_argument('--log-level=3')
    chrome_option.add_argument('--start-maximized')
    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome(options=chrome_option)

    tku_coord = [25.174542, 121.450259] 
    map_for_students = folium.Map(location=tku_coord, zoom_start=16, tiles="OpenStreetMap")
    map_for_teachers = folium.Map(location=tku_coord, zoom_start=16, tiles="OpenStreetMap")

    for index, row in filtered_data.iterrows():
        address = row[title[4]]
        property_label = row[title[5]]
        property_type = row[title[6]]
        self_safety_check = row[title[7]]

        coord = get_lat_lng(driver, address)
        sleep(1)

        if coord:
            icon_color = {
                "大樓 building": 'blue',
                "公寓 apartment": 'orange',
                "平房 bungalow": 'purple',
                "我的租屋處是安全的，不需要教官到場訪視 My rental place is safe and there is no need for instructors to visit the place": 'green',
                "我的租屋處有些許不安全，但我可以自己處理並主動回報，不需要教官到場訪視 My rental apartment is a little unsafe, but I can handle it myself and report it proactively. I don’t need an instructor to visit the place": 'orange',
                "需要教官到我的租屋處再幫忙檢視 I need an instructor to come to my rental office and check it again": 'red'
            }

            if property_label in ["無", "none", "None"]:
                folium.Marker(
                    location=coord, 
                    icon=folium.Icon(color=icon_color[property_type])
                ).add_to(map_for_students)
                folium.Marker(
                    location=coord, 
                    icon=folium.Icon(color=icon_color[self_safety_check])
                ).add_to(map_for_teachers)
            else:
                folium.Marker(
                    location=coord, 
                    popup=property_label,
                    icon=folium.Icon(color=icon_color[property_type])
                ).add_to(map_for_students)
                folium.Marker(
                    location=coord, 
                    popup=property_label,
                    icon=folium.Icon(color=icon_color[self_safety_check])
                ).add_to(map_for_teachers)
        else:
            print(f'{address} coordinates not found!')

    map_for_students.save('map_for_students.html')
    print("地圖已保存為 'map_for_students.html'")

    map_for_teachers.save('map_for_teachers.html')
    print("地圖已保存為 'map_for_teachers.html'")
    