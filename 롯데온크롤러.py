from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PySide6.QtWidgets import QFileDialog
import json

import re
import time
import pyperclip
import pandas as pd
from openpyxl import Workbook

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QCheckBox, QVBoxLayout, QHBoxLayout, QTextEdit, QGridLayout
)

#캐시디렉토리는 배포때 제거, Key.command 명령어 수정필요 윈도우용으로
caches_dir = "/Users/iseung-ug/Library/Application Support/Code"
new_Caches = "user_caches"

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
    
        self.driver = None  # Class-level WebDriver instance
        self.setup_ui()

    def setup_ui(self):

        self.setWindowTitle("롯데ON 데이터 수집기")
        self.setGeometry(100, 100, 800, 600)

        # 메인 위젯 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # 레이아웃 설정
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # 상단 타이틀 라벨
        title_label = QLabel("  롯데ON 데이터 수집기")
        title_label.setStyleSheet("background-color: red; color: white; font-size: 18px; font-weight: bold; border: 2px solid black; border-radius: 5px;")
        title_label.setFixedHeight(30)
        main_layout.addWidget(title_label)

        # ID와 PW 입력 섹션
        id_pw_layout = QHBoxLayout()
        self.id_input = QLineEdit()
        self.pw_input = QLineEdit()
        self.pw_input.setEchoMode(QLineEdit.Password)  # 비밀번호 입력 필드 숨김

        id_label = QLabel("ID")
        pw_label = QLabel("PW")
        login_button = QPushButton("카카오 로그인")
        lotte_login_button = QPushButton("롯데온 접속")
        login_button.clicked.connect(self.handle_login)  # 이벤트 연결
        lotte_login_button.clicked.connect(self.lotte_login)
        login_button.setStyleSheet("background-color: lightgray; font-weight: bold;")
        lotte_login_button.setStyleSheet("background-color: red; font-weight: bold;")

        id_pw_layout.addWidget(id_label)
        id_pw_layout.addWidget(self.id_input)
        id_pw_layout.addWidget(pw_label)
        id_pw_layout.addWidget(self.pw_input)
        id_pw_layout.addWidget(login_button)
        id_pw_layout.addWidget(lotte_login_button)

        main_layout.addLayout(id_pw_layout)

        # 데이터 수집 섹션
        collection_layout = QHBoxLayout()

        multi_page_label = QLabel("페이지 수")
        self.multi_page_input = QLineEdit()
        collection_layout.addWidget(multi_page_label)
        collection_layout.addWidget(self.multi_page_input)

        main_layout.addLayout(collection_layout)

        # 체크박스 섹션
        checkbox_layout = QHBoxLayout()

        self.address_checkbox = QCheckBox("성함")
        self.address_checkbox.setChecked(True)
        self.name_checkbox = QCheckBox("주소")
        self.name_checkbox.setChecked(True)
        self.shipment_checkbox = QCheckBox("송장번호")
        self.shipment_checkbox.setChecked(True)
        self.company_checkbox = QCheckBox("택배사")
        self.company_checkbox.setChecked(True)
        self.state_checkbox = QCheckBox("배송상태")
        self.state_checkbox.setChecked(True)

        checkbox_layout.addWidget(self.address_checkbox)
        checkbox_layout.addWidget(self.name_checkbox)
        checkbox_layout.addWidget(self.shipment_checkbox)
        checkbox_layout.addWidget(self.company_checkbox)
        checkbox_layout.addWidget(self.state_checkbox)

        main_layout.addLayout(checkbox_layout)

        # 수집 버튼 섹션
        button_layout = QGridLayout()
        collect_button = QPushButton("수 집")
        collect_button.setStyleSheet("background-color: lightgreen; font-weight: bold;")
        collect_button.clicked.connect(self.handle_collect)
        button_layout.addWidget(collect_button)
        main_layout.addLayout(button_layout)

        # 텍스트 에리어 섹션
        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("결과화면")
        main_layout.addWidget(self.text_area)

        # 하단 버튼
        bottom_button = QPushButton("다운로드")
        bottom_button.setStyleSheet("background-color: lightgreen; font-weight: bold;")
        bottom_button.clicked.connect(self.download_results)
        main_layout.addWidget(bottom_button)

    def initialize_driver(self):
        useragent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        try:
            # 드라이버가 None이거나 세션이 종료된 상태인지 확인
            if self.driver is None or not self.is_driver_alive():
                options = Options()
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_argument(f"user-agent={useragent}")
                #options.add_argument(f"--user-data-dir={new_Caches}")
                options.add_argument("--profile-directory=Default")
                options.add_experimental_option("detach", True)
                #service = Service("/Users/iseung-ug/.cache/selenium/chromedriver/mac-arm64/131.0.6778.264/chromedriver")
                self.driver = webdriver.Chrome(options=options)
                self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            return self.driver
        except Exception as e:
            self.text_area.append(f"드라이버 초기화 중 오류 발생: {str(e)}")
            return None
        
    def is_driver_alive(self):
        """드라이버가 활성 상태인지 확인하는 메서드"""
        if self.driver is None:
            return False
        try:
            # 간단한 명령을 실행해서 드라이버가 응답하는지 확인
            self.driver.current_url
            return True
        except:
            return False

    def safe_quit_driver(self):
        """안전하게 드라이버를 종료하는 메서드"""
        try:
            if self.is_driver_alive():
                self.driver.quit()
        except Exception as e:
            print(f"드라이버 종료 중 오류 발생: {str(e)}")
        finally:
            self.driver = None

    def handle_login(self):
        try:
            self.headless = False
            self.initialize_driver()
            if self.driver:
                user_id = self.id_input.text()
                user_pw = self.pw_input.text()
                self.run_selenium_login(user_id, user_pw)
        except Exception as e:
            self.text_area.append(f"로그인 처리 중 오류 발생: {str(e)}")

    def handle_collect(self):
        try:
            if self.driver:
                stop_page = int(self.multi_page_input.text()) if self.multi_page_input.text().isdigit() else 1
                self.text_area.append("데이터 수집 시작...")
                self.run_selenium_collection(stop_page)
                self.driver.quit()
        except Exception as e:
            self.text_area.append(f"데이터 수집 중 오류 발생: {str(e)}")

    def run_selenium_login(self, user_id, user_pw):
        self.driver.get("https://www.lotteon.com")
        
        kakao_url = "https://accounts.kakao.com/login/?continue=https%3A%2F%2Fkauth.kakao.com%2Foauth%2Fauthorize%3Fproxy%3DeasyXDM_Kakao_zekzs7gq7co_provider%26ka%3Dsdk%252F1.43.5%2520os%252Fjavascript%2520sdk_type%252Fjavascript%2520lang%252Fko-KR%2520device%252FMacIntel%2520origin%252Fhttps%25253A%25252F%25252Fwww.lotteon.com%26origin%3Dhttps%253A%252F%252Fwww.lotteon.com%26response_type%3Dcode%26redirect_uri%3Dkakaojs%26state%3D45r6e5vfal7kunhmjerlu%26client_id%3Ddcc55d89bf71280ca514cc30d7e6dc32%26through_account%3Dtrue&talk_login=hidden#login"

        self.driver.get(kakao_url)

        id_selector = "#loginId--1"
        pw_selector = "#password--2"
        time.sleep(2)

        id_input = self.driver.find_element(By.CSS_SELECTOR, id_selector)
        id_input.click()

        time.sleep(1)

        pyperclip.copy(user_id)
        ActionChains(self.driver).key_down(Keys.COMMAND).send_keys("v").key_down(Keys.COMMAND).perform()

        time.sleep(1)

        pw_input = self.driver.find_element(By.CSS_SELECTOR, pw_selector)
        pw_input.click()

        time.sleep(1)

        pyperclip.copy(user_pw)
        ActionChains(self.driver).key_down(Keys.COMMAND).send_keys("v").key_down(Keys.COMMAND).perform()

        time.sleep(2)

        login_selector = "#mainContent > div > div > form > div.confirm_btn > button.btn_g.highlight.submit"
        login = self.driver.find_element(By.CSS_SELECTOR, login_selector)
        login.click()
        time.sleep(1)
        url = "https://www.lotteon.com/p/member/login/common/"
        self.driver.get(url)
        kakao_class = "kakaoLoginBtn"
        kakao_login = self.driver.find_element(By.CLASS_NAME, kakao_class)
        kakao_login.click()
        time.sleep(4)
        
    def lotte_login(self):


        if self.driver:
            self.cookies = self.driver.get_cookies()
            
            with open("cookies.json", "w") as f:
                json.dump(self.cookies, f)
            self.driver.quit()
       
        self.getNewDriver()

    def getNewDriver(self):
        
        with open("cookies.json", "r") as f:
            cookies = json.load(f)
        user_id = self.id_input.text()
        user_pw = self.pw_input.text()
        caches_dir = "/Users/iseung-ug/Library/Application Support/Code"
        useragent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        options = Options()
        #options.add_argument(f"--user-data-dir={new_Caches}")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument(f"user-agent={useragent}")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--window-size=1920,1080")
        #options.add_argument("--headless")
        
        options.add_experimental_option("detach", True)
        #service = Service("/Users/iseung-ug/.cache/selenium/chromedriver/mac-arm64/131.0.6778.264/chromedriver")
        self.driver = webdriver.Chrome(options=options)

        
        self.driver.implicitly_wait(10)
        self.driver.get("https://www.lotteon.com")
        for cookie in cookies:
            self.driver.add_cookie(cookie)

        self.driver.refresh()    
        kakao_url = "https://accounts.kakao.com/login/?continue=https%3A%2F%2Fkauth.kakao.com%2Foauth%2Fauthorize%3Fproxy%3DeasyXDM_Kakao_zekzs7gq7co_provider%26ka%3Dsdk%252F1.43.5%2520os%252Fjavascript%2520sdk_type%252Fjavascript%2520lang%252Fko-KR%2520device%252FMacIntel%2520origin%252Fhttps%25253A%25252F%25252Fwww.lotteon.com%26origin%3Dhttps%253A%252F%252Fwww.lotteon.com%26response_type%3Dcode%26redirect_uri%3Dkakaojs%26state%3D45r6e5vfal7kunhmjerlu%26client_id%3Ddcc55d89bf71280ca514cc30d7e6dc32%26through_account%3Dtrue&talk_login=hidden#login"
        
        self.driver.execute_script("window.open('https://www.lotteon.com', '_blank');")
        
        time.sleep(2)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[1])

        orders_selector = "#mainLayout > header > div > div.util.innerContent > div > ul > li:nth-child(3) > a"
        orders = self.driver.find_element(By.CSS_SELECTOR, orders_selector)
        self.driver.execute_script("arguments[0].click()", orders)
        self.text_area.setText("접속성공")
        
        print("메인 주문관리 탭으로 이동")


    def run_selenium_collection(self, stop_page):

        tabs = self.driver.window_handles
        #self.driver.switch_to.window(tabs[0])
        self.driver.execute_script("window.open('https://www.lotteon.com', '_blank');")
        
        time.sleep(2)
        tabs = self.driver.window_handles
        self.driver.switch_to.window(tabs[-1])

        orders_selector = "#mainLayout > header > div > div.util.innerContent > div > ul > li:nth-child(3) > a"
        orders = self.driver.find_element(By.CSS_SELECTOR, orders_selector)
        self.driver.execute_script("arguments[0].click()", orders)
        self.text_area.setText("수집시작")
        time.sleep(3)
        input()
        i = 3
        next_page = 1
        results = []

        while i < 3 + stop_page * 15:
            time.sleep(2)
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[1])
            print("메인 주문관리 탭으로 이동")

            try:
                list_selector = "#content > div > div.contentWrap > div.myLotteWrap > div"
                list_element = self.driver.find_element(By.CSS_SELECTOR, list_selector)
                product_element = list_element.find_element(By.XPATH, f"./*[{i}]")

                if product_element.get_attribute("class") == "btnCenter":
                    next_page += 1
                    load_page_element = product_element.find_element(By.XPATH, "./*[1]")
                    self.driver.execute_script("arguments[0].click();", load_page_element)
                    print(f"'더보기' 버튼 클릭: {i}")
                    time.sleep(2)

            except Exception as e:
                print("리스트 탐색 중 오류:", e)

            orders_url = self.driver.current_url
            self.driver.execute_script(f"window.open('{orders_url}', '_blank');")
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[2])
            if next_page > 1:
                loading_count = 1
                while loading_count < next_page:
                    loading_option = (loading_count)*15+3
                    try:
                        # 리스트 탐색
                        time.sleep(1)
                        list_selector = "#content > div > div.contentWrap > div.myLotteWrap > div"
                        list_element = self.driver.find_element(By.CSS_SELECTOR, list_selector)
                        product_element = list_element.find_element(By.XPATH, f"./*[{loading_option}]")
                        load_page_element = product_element.find_element(By.XPATH, "./*[1]")
                        self.driver.execute_script("arguments[0].click();", load_page_element)
                        print("새 탭 페이지 넘김 성공")
                        loading_count += 1
                    except:
                        print("새 탭에서 페이지넘김 오류")
                        time.sleep(1)

            try:
                time.sleep(2)
                list_selector = "#content > div > div.contentWrap > div.myLotteWrap > div"
                list_element = self.driver.find_element(By.CSS_SELECTOR, list_selector)
                product_element = list_element.find_element(By.XPATH, f"./*[{i}]")

                infoto_selector = ".topInformation.grayBox .buttons button"
                infoto = product_element.find_element(By.CSS_SELECTOR, infoto_selector)
                self.driver.execute_script("arguments[0].click();", infoto)
                

                        # 주문 상태, 고객 정보 등 추출
                #주문상태
                try:
                    time.sleep(1)
                    order_status_class = "status"
                    order_state_element = self.driver.find_element(By.CLASS_NAME, order_status_class)
                    # JavaScript로 텍스트 추출
                    order_state = self.driver.execute_script("return arguments[0].textContent;", order_state_element)
                    order_state = str.strip(order_state)
                    print(order_state)
                except Exception as e:
                    print("리로딩")
                    time.sleep(3)

                post_no, post_company = None, None

#content > div > div.contentWrap > div.myLotteWrap > div > div > div.orderGroupWrap > div.orderListWrap.RoundType > div.orderGoodsItem > div.orderStatusInfo > div.statusInfoWrap > span.status


                       #운송장번호
                if order_state == "배송진행중" or order_state == "배송중":
                    print("운송장번호 로직 구현")
                    try:
                        post_btn_class = "orderStatusInfoButtons"
                        post_btn_element = self.driver.find_element(By.CLASS_NAME, post_btn_class).find_element(By.XPATH, "./*[1]")
                        self.driver.execute_script("arguments[0].click();", post_btn_element)
                        time.sleep(1)
                        #input("체킹중")
                    except:
                        print("배송상세정보 찾기 에러")

                    try:
                        post_info_class = "deliveryTrackingInfo"
                        post_info_element = self.driver.find_element(By.CLASS_NAME, post_info_class)
                    except:
                        print("페이지로딩 실패")

                    try:    
                        post_info_element = post_info_element.find_element(By.XPATH, "./*[2]")
                        post_no_element = post_info_element.find_element(By.XPATH, "./*[4]")
                        post_no = self.driver.execute_script("return arguments[0].textContent;", post_no_element)
                        post_company_element = post_info_element.find_element(By.CLASS_NAME, "show-call-icon")
                        post_company = self.driver.execute_script("return arguments[0].textContent;", post_company_element)
                        print(post_company + " " + post_no)
                    except:
                        print("데이터수집실패")

                if order_state == "배송완료":
                    try:
                        post_btn_class = "orderStatusInfoButtons"
                        post_btn_element = self.driver.find_element(By.CLASS_NAME, post_btn_class).find_element(By.XPATH, "./*[4]")
                        self.driver.execute_script("arguments[0].click();", post_btn_element)
                        time.sleep(1)
                        #input("체킹중")
                    except:
                        print("배송상세정보 찾기 에러")

                    try:
                        post_info_class = "deliveryTrackingInfo"
                        post_info_element = self.driver.find_element(By.CLASS_NAME, post_info_class)
                    except:
                        print("페이지로딩 실패")

                    try:    
                        post_info_element = post_info_element.find_element(By.XPATH, "./*[2]")
                        post_no_element = post_info_element.find_element(By.XPATH, "./*[4]")
                        post_no = self.driver.execute_script("return arguments[0].textContent;", post_no_element)
                        post_company_element = post_info_element.find_element(By.CLASS_NAME, "show-call-icon")
                        post_company = self.driver.execute_script("return arguments[0].textContent;", post_company_element)
                        print(post_company + " " + post_no)
                    except:
                        print("데이터수집실패")

                #고객명
                list = self.driver.find_element(By.CLASS_NAME, "informationArea").find_element(By.CLASS_NAME, "list")
                
                #content > div > div.contentWrap > div.myLotteWrap > div > div > div.orderGroupWrap > div.informationArea > div.infoArea > div > ul > li:nth-child(1) > div.text
                cust_name_element = list.find_element(By.XPATH, "./*[1]/*[1]/*[2]")
                cust_name = self.driver.execute_script("return arguments[0].textContent;", cust_name_element)
                cust_name = str.strip(cust_name)
                print(cust_name)

                #고객주소
                cust_address_selector = "#content > div > div.contentWrap > div.myLotteWrap > div:nth-child(2) > div > div.orderGroupWrap > div.informationArea > div.infoArea > div > ul > li:nth-child(3) > div.text"
                cust_address_element = list.find_element(By.XPATH, "./*[1]/*[3]/*[2]")
                cust_address = self.driver.execute_script("return arguments[0].textContent;", cust_address_element)
                cust_address = re.sub(r'\s{3,}', ' ', str.strip(cust_address))
                print(cust_address)
                print("\n")

                result = {}
                
                if self.name_checkbox.isChecked:
                    result["성함"] = cust_name
                if self.address_checkbox.isChecked:
                    result["주소"] = cust_address
                if self.shipment_checkbox.isChecked:
                    result["송장번호"] = post_no
                if self.company_checkbox.isChecked:
                    result["택배사"] = post_company
                if self.state_checkbox.isChecked:
                    result["배송상태"] = order_state

                
                results.append(result)
                self.text_area.append(str(result))

                i += 1
            except Exception as e:
                print("상세 정보 추출 중 오류:", e)

            self.driver.close()
            tabs = self.driver.window_handles
            self.driver.switch_to.window(tabs[1])

        self.results = results

    def closeEvent(self, event):
            """앱 종료 시 실행되는 메서드"""
            self.safe_quit_driver()
            event.accept()

    def download_results(self):
        # 결과가 없을 경우 메시지 출력
        if not hasattr(self, 'results') or not self.results:
            self.text_area.append("결과가 없습니다.")
            return

        # 워크북 생성
        wb = Workbook()
        ws = wb.active
        ws.title = "수집 결과"

        # 헤더 추가
        headers = ["성함", "주소", "송장번호", "택배사", "배송상태"]
        ws.append(headers)

        # 데이터 추가
        for result in self.results:
            ws.append([result["성함"], result["주소"], result["송장번호"], result["택배사"], result["배송상태"]])

        # 열 너비 자동 조정
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # 열 이름 가져오기
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception as e:
                    print(f"셀 값 처리 중 오류: {e}")
            adjusted_width = max_length + 2  # 약간의 여유 공간 추가
            ws.column_dimensions[column].width = adjusted_width

        # 저장 경로 설정
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "엑셀 파일 저장",
            "",
            "Excel Files (*.xlsx);;All Files (*)",
        )

        if save_path:
            try:
                # 파일 확장자가 없으면 추가
                if not save_path.endswith(".xlsx"):
                    save_path += ".xlsx"
                wb.save(save_path)
                self.text_area.append(f"엑셀 파일로 저장되었습니다: {save_path}")
            except Exception as e:
                self.text_area.append(f"파일 저장 중 오류 발생: {e}")
        else:
            self.text_area.append("저장이 취소되었습니다.")


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
    
    