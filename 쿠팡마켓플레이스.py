from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import requests, time, random

def random_sleep():
    randnum = random.randint(3, 10) / 10
    time.sleep(randnum)
    

data = [ 
80987, 
81005
]
length = len(data) - 1

# 사용자 지정 프로필 경로 (캐시 저장 위치)
user_data_dir = "/Users/iseung-ug/Desktop/보문산보물상점/소스코드/Project/basic_selenium/sel1/coopangwing_caches"
user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36'


# Chrome 옵션 설정
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument(f"user-data-dir={user_data_dir}")  # 사용자 데이터 저장 경로
chrome_options.add_argument("profile-directory=default")  # 기본 프로필 사용 (선택 사항)
chrome_options.add_argument(f"--user-agent={user_agent}")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_argument("--disable-blink-features=AutomationControlled")


# ChromeDriver 실행
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

trends_search_url = "https://wing.coupang.com/tenants/rfm-ss/coupang-trends/popularity-search"

category_btn = False

def category_leaf_or_full():
    return True

def reload():
    current_window = driver.window_handles
    driver.switch_to.window(current_window)

# 웹 페이지 열기
driver.get(trends_search_url)
input("로그인 완료?")

#요소 셀렉터
name_selector = "div._product-info-container_1g00b_39 > div._details_1g00b_45 > div._details-left_1g00b_51 > div._subject_1g00b_29._clickable_1g00b_100 > span"
price_selector = "div._product-info-container_1g00b_39 > div._emphasize-wrapper_1g00b_63 > div:nth-child(3) > div > strong"
rating_selector = "div._product-info-container_1g00b_39 > div._emphasize-wrapper_1g00b_63 > div:nth-child(1) > div > div > strong"
ratings_selector = "div._product-info-container_1g00b_39 > div._emphasize-wrapper_1g00b_63 > div:nth-child(2) > div > strong"
views_selector = "div._product-info-container_1g00b_39 > div._emphasize-wrapper_1g00b_63 > div:nth-child(4) > div > span"

datalist = []
for categoryCode in data:
    driver.get(f"https://wing.coupang.com/tenants/rfm-ss/coupang-trends/popularity-search?categoryCode={categoryCode}")

    #리스트 요소 찾기
    loof = True
    count = 1
    while loof:
        try:
            time.sleep(3)
            products = driver.find_elements(By.CLASS_NAME, "_product_card_1g00b_1")
            loof = False
            count += 1
        except:
            if count == 4:
                print("리스트 요소 검색 실패")
                break

    #리스트 요소 순회
    for product in products:
        try:
            #카테고리 leaf 노드 검색
            parent_element = product.find_element(By.CSS_SELECTOR, "div._product-info-container_1g00b_39 > div._details_1g00b_45 > div._details-left_1g00b_51 > div._container_1ml8k_1")
            child_elements = parent_element.find_elements(By.XPATH, "./*")
            last_child = child_elements[-1]
            category_leaf = driver.execute_script("return arguments[0].textContent;", last_child)
            full_category = ""
            for childs in child_elements:
                temp = ">" + driver.execute_script("return arguments[0].textContent;", childs)
                full_category = full_category + temp 

            #full_category = driver.execute_script("return arguments[0].textContent;", parent_element)

            #제품이름 요소 검색
            product_name_e = product.find_element(By.CSS_SELECTOR, name_selector)
            product_name = driver.execute_script("return arguments[0].textContent;", product_name_e)

            #제품이름 링크 추출
            driver.execute_script("arguments[0].click();", product_name_e)
            windows = driver.window_handles
            #링크 탭 전환
            driver.switch_to.window(windows[1])
            product_url = driver.current_url
            time.sleep(0.6)
            driver.close()
            #메인 탭 전환
            windows = driver.window_handles
            driver.switch_to.window(windows[0])

            #제품 가격, 별점, 리뷰수, 조회수 요소 검색
            product_price_e = product.find_element(By.CSS_SELECTOR, price_selector)
            product_rating_e = product.find_element(By.CSS_SELECTOR, rating_selector)
            product_ratings_e = product.find_element(By.CSS_SELECTOR, ratings_selector)
            product_views_e = product.find_element(By.CSS_SELECTOR, views_selector)

            #값 추출
            product_price = driver.execute_script("return arguments[0].textContent;", product_price_e)
            product_rating = driver.execute_script("return arguments[0].textContent;", product_rating_e)
            product_ratings = driver.execute_script("return arguments[0].textContent;", product_ratings_e)
            product_views = driver.execute_script("return arguments[0].textContent;", product_views_e)
            
            #출력
            print(full_category)
            print(category_leaf)
            print(product_name)
            print(product_url)
            print(product_price)
            print(product_rating)
            print(product_ratings)
            print(product_views)

            #카테고리 전부 출력할지 선택
            if category_btn:
                real_category = category_leaf
            else:
                real_category = full_category

            #조회수 분류
            if product_views == "1,000회 미만":
                minparts = 0
                maxparts = 1000
            elif product_views == "집계중":
                minparts = 0
                maxparts = 0
            else:
                parts = [part.strip() for part in product_views.split("-")]
                if "만" in parts[0]:
                    minparts = int(parts[0].replace("만", "")) * 10000
                    print(minparts)
                elif "십만" in parts[0]:
                    minparts = int(parts[0].replace("십만", "")) *100000
                    print(minparts)
                else:
                    minparts = int(parts[0].replace(",", ""))
                    print(minparts)

                if "만" in parts[1]:
                    maxparts = int(parts[1].replace("만회", ""))*10000
                    print(maxparts)
                elif "십만" in parts[1]:
                    maxparts = int(parts[1].replace("십만회", "")) *100000
                    print(maxparts)
                else:
                    parts
                    maxparts = parts[1].replace("회", "")
                    maxparts = (maxparts.replace(",", ""))
                    print(maxparts)

            

            #데이터 저장
            datalist.append({"카테고리" : real_category, "제품명" : product_name, "링크" : product_url, "판매가격" : product_price, "리뷰별점" : product_rating, "리뷰수" : product_ratings, "최소조회수" : minparts, "최대조회수" : maxparts})

            #0.3~1초 대기
            random_sleep()
        except Exception as e:
            print(f"에러내용:{e}")
            break

        time.sleep(5) 


import pandas as pd
import openpyxl

# 데이터 저장할 엑셀 파일 이름 설정
excel_filename = "coupang_products.xlsx"

# 데이터프레임 변환
df = pd.DataFrame(datalist)

# 엑셀로 저장
df.to_excel(excel_filename, index=False, encoding="utf-8-sig", engine="openpyxl")

print(f"데이터가 '{excel_filename}' 파일로 저장되었습니다.")




