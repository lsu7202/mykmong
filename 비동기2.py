import multiprocessing
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
import os
import asyncio
import concurrent.futures
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor

yahoo = "https://map.yahoo.co.jp/address"
query = "コインランドリー"
searchYahoo = f"https://map.yahoo.co.jp/search?q={query}"

# 비동기 태스크를 위한 함수
async def async_process_task(idx):
    # 프로세스 풀에서 실행할 함수
    loop = asyncio.get_event_loop()
    with ProcessPoolExecutor(max_workers=1) as executor:
        return await loop.run_in_executor(executor, process_task, idx)

def process_task(idx):
    # 실제 작업을 수행하는 함수
    process = scrapyProcess(idx)
    process.getDriver(idx)
    return f"Process {idx} completed"

class scrapyProcess(multiprocessing.Process):
    def __init__(self, idx):
        multiprocessing.Process.__init__(self)
        self.idx = idx
        self.driver = None
        
    def getDriver(self, idx):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--headless=new")
        options.add_experimental_option("detach", True)

        self.driver = webdriver.Chrome(options=options)
        self.driver.implicitly_wait(10)
        self.driver.get(yahoo)
        #input("get!")
        
        headerId = []
        for first in self.driver.find_elements(By.CLASS_NAME, "SearchAddressResults__list"):
            headerId.append(first.get_attribute("id"))

        heads = []
        for Id in headerId:
            for head in self.driver.find_elements(By.CSS_SELECTOR, f"#{Id} > ul > *"):
                heads.append(head)

        goDodobuyun = heads[idx].find_element(By.XPATH, "./*[1]")
        self.dodobuhyunName = self.driver.execute_script("return arguments[0].textContent;", goDodobuyun)
        self.driver.execute_script("arguments[0].click()", goDodobuyun)
        

        self.getDetails()
        for idx in range(self.Size):
            self.dodobuhyunUrl = self.driver.current_url
            self.getDetails()
            self.detail_text = self.driver.execute_script("return arguments[0].textContent", self.detailBtns[idx])
            self.driver.execute_script("arguments[0].click()", self.detailBtns[idx])
            print(f"{self.dodobuhyunName} : {idx}/{self.Size}")
            self.getDatas()
            self.BackToDodoBuyun()


    def getTitle(self):
        self.driver.implicitly_wait(10)
        Title = self.driver.find_element(By.CSS_SELECTOR, "#search_address_page > div.SearchAddressResults__currentPosition > div > font > font")
        Title = self.driver.execute_script("return arguments[0].textContent;", Title)
        return Title
    
    def getDetails(self):
        self.driver.implicitly_wait(10)
        headerId = []
        for first in self.driver.find_elements(By.CLASS_NAME, "SearchAddressResults__list"):
            headerId.append(first.get_attribute("id"))

        self.detailBtns = []
        for Id in headerId:
            for head in self.driver.find_elements(By.CSS_SELECTOR, f"#{Id} > ul > *"):
                self.detailBtns.append(head.find_element(By.XPATH, "./*[1]"))
        
        self.Size = len(self.detailBtns)

    def getDatas(self):
        driver = self.driver
        driver.implicitly_wait(10)
        driver.get(searchYahoo)
        print("크롤링작업")
        while True:
            try:
                data = []
                lists_parent = driver.find_element(By.CSS_SELECTOR, "#search_keyword > div.SearchKeywordResults > ul")

                lists = lists_parent.find_elements(By.CLASS_NAME, "SearchKeywordResults__listItem")

                for i in lists:
                    "/html/body/div/div/div[2]/div[2]/div[1]/div[2]/div[2]/ul/li[1]/button/div/div/div[1]"

                    url_btn = i.find_element(By.CLASS_NAME, "SearchKeywordResults__listItemButton")   


                    store_name = i.find_element(By.XPATH, "./*[1]/*[2]/*[1]/*[1]")
                    store_name = driver.execute_script("return arguments[0].textContent;", store_name)


                    #페이지 이동 없이 url 추출
                    try:
                        store_url = driver.execute_script("""
                            let targetUrl = null;
                            let originalOpen = window.open;

                            // window.open을 가로채서 URL 저장
                            window.open = function(url) {
                                targetUrl = url;
                                return null;
                            };

                            arguments[0].click();  // 클릭 실행

                            return targetUrl;  // 이동될 URL 반환
                        """, url_btn)
                        #iframe = driver.find_element(By.CLASS_NAME, "POI__externalContentFrame")
                        # iframe 찾기 (숨겨진 상태)

                        iframe = driver.find_element(By.TAG_NAME, "iframe")

                        # JavaScript로 iframe을 보이게 만들기
                        driver.execute_script("arguments[0].style.display = 'block'; arguments[0].width = '600'; arguments[0].height = '400';", iframe)

                        # 이제 iframe으로 전환
                        driver.switch_to.frame(iframe)
            
                        try:
                            address = driver.find_element(By.CLASS_NAME, "AddressModule_AddressModule__contentsBodyDetailsAddress__d52ba")
                            address = driver.execute_script("return arguments[0].textContent;", address)
                        except:
                            address = ""
                        
                        try:
                            ratings = driver.find_element(By.CLASS_NAME, "SummaryPlace_SummaryPlace__informationOverviewSummaryEvaluation__bvOXM")
                            ratings = driver.execute_script("return arguments[0].textContent;", ratings)
                        except:
                            ratings = ""

                        windows = driver.window_handles
                        driver.switch_to.window(windows[0])

                        # 데이터 저장
                        data.append([self.detail_text, store_name, address, ratings])
                        

                        #print(driver.page_source)
                        

                    except Exception as e:
                        print(e)
            except:
                input("리스트찾기에러")
                self.getDatas()

                # DataFrame 생성 및 엑셀로 저장
            file_path = f"{self.dodobuhyunName}.xlsx"
        
            if os.path.exists(file_path):
                existing_df = pd.read_excel(file_path, engine='openpyxl')
                new_df = pd.DataFrame(data, columns=["detail", "Store Name", "Address", "Ratings"])
                df = pd.concat([existing_df, new_df], ignore_index=True)
            else:
                df = pd.DataFrame(data, columns=["detail", "Store Name", "Address", "Ratings"])
            df.to_excel(file_path, index=False)
            print(f"엑셀 저장 완료: {file_path}")


            nextBtn = driver.find_element(By.CSS_SELECTOR, "#search_keyword > div.SearchKeywordResults > div > div > div.Pagination__next > button")
            if nextBtn.get_attribute("class") == "Pagination__controllerButton Pagination__controllerButton--disabled":
                print("다음페이지로 이동")
                break
            else:
                driver.execute_script("arguments[0].click();", nextBtn)
                print(nextBtn.get_attribute("class"))

            


    def BackToDodoBuyun(self):
        self.driver.implicitly_wait(10)
        self.driver.get(self.dodobuhyunUrl)
         

    def run(self):
        idx = self.idx
        self.getDriver(idx)
     

# 비동기 메인 함수
async def main():
    # 동시에 실행할 프로세스의 수
    num_processes = 5
    tasks = []
    
    # 비동기 작업 생성
    for idx in [43,15,24,30,16]: #26 27 28 29 30 미완
        task = asyncio.create_task(async_process_task(idx))
        tasks.append(task)
    
    # 모든 작업 완료 대기
    results = await asyncio.gather(*tasks)
    print("모든 프로세스 완료:", results)

if __name__ == "__main__":
    # multiprocessing 시작 방식 설정
    multiprocessing.set_start_method('spawn', force=True)
    
    # 메인 비동기 루프 실행
    asyncio.run(main())