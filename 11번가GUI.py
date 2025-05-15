from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PySide6.QtWidgets import QFileDialog

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
def get_driver():
    options = Options()
    options.add_argument(f"--user-data-dir={caches_dir}")
    options.add_argument("--profile-directory=Default")
    options.add_experimental_option("detach", True)
    service = Service("/Users/iseung-ug/.cache/selenium/chromedriver/mac-arm64/131.0.6778.264/chromedriver")  # chromedriver 경로 지정
    driver = webdriver.Chrome(service=service, options=options)
    return driver


class MainWindow(QMainWindow):

    def __init__(self):
        
        super().__init__()
        self.driver = None  # Class-level WebDriver instance
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("11번가 데이터 수집기")
        self.setGeometry(100, 100, 800, 600)

        # 메인 위젯 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 레이아웃 설정
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # 상단 타이틀 라벨
        title_label = QLabel("  11번가 데이터 수집기")
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
        market_login_button = QPushButton("11번가 접속")
        login_button.clicked.connect(self.handle_login)  # 이벤트 연결
        market_login_button.clicked.connect(self.market_login)
        login_button.setStyleSheet("background-color: yellow; color: black; font-weight: bold;")

        market_login_button.setStyleSheet("background-color: red; color: white; font-weight: bold;")

        id_pw_layout.addWidget(id_label)
        id_pw_layout.addWidget(self.id_input)
        id_pw_layout.addWidget(pw_label)
        id_pw_layout.addWidget(self.pw_input)
        id_pw_layout.addWidget(login_button)
        id_pw_layout.addWidget(market_login_button)

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
        try:
            # 드라이버가 None이거나 세션이 종료된 상태인지 확인
            if self.driver is None or not self.is_driver_alive():
                options = Options()
                options.add_argument(f"--user-data-dir={caches_dir}")
                options.add_argument("--profile-directory=Default")
                options.add_experimental_option("detach", True)
                service = Service("/Users/iseung-ug/.cache/selenium/chromedriver/mac-arm64/131.0.6778.264/chromedriver")
                self.driver = webdriver.Chrome(service=service, options=options)
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
            self.initialize_driver()
            if self.driver:
                user_id = self.id_input.text()
                user_pw = self.pw_input.text()
                self.run_selenium_login(user_id, user_pw)
        except Exception as e:
            self.text_area.append(f"로그인 처리 중 오류 발생: {str(e)}")

    def handle_collect(self):
        try:
            self.initialize_driver()
            if self.driver:
                stop_page = int(self.multi_page_input.text()) if self.multi_page_input.text().isdigit() else 1
                self.text_area.append("데이터 수집 시작...")
                self.run_selenium_collection(stop_page)
        except Exception as e:
            self.text_area.append(f"데이터 수집 중 오류 발생: {str(e)}")

    def market_login(self):
        try:
            self.initialize_driver()
            if self.driver:
                url = "https://login.11st.co.kr/"
                self.driver.get(url)
                kakao_selector = "body > div.login-wrapper > div:nth-child(2) > div:nth-child(3) > div > ul > li:nth-child(1) > a > img"
                kakao_login = self.driver.find_element(By.CSS_SELECTOR, kakao_selector)
                kakao_login.click()
                time.sleep(4)

                # 탭 처리
                self.handle_tabs_after_login()
        except Exception as e:
            self.text_area.append(f"마켓 로그인 중 오류 발생: {str(e)}")

    def handle_tabs_after_login(self):
        """로그인 후 탭 처리를 위한 메서드"""
        try:
            tabs = self.driver.window_handles
            if len(tabs) > 1:
                self.driver.switch_to.window(tabs[1])
                self.driver.close()
                self.driver.switch_to.window(tabs[0])

            orders_selector = "#gnb > div > div.b_header_gnb > div > div.c-util-user > ul > li.order > a"
            orders = self.driver.find_element(By.CSS_SELECTOR, orders_selector)
            orders.click()
            time.sleep(3)
            
            tabs = self.driver.window_handles
            if len(tabs) > 1:
                self.driver.switch_to.window(tabs[1])
        except Exception as e:
            raise Exception(f"탭 처리 중 오류 발생: {str(e)}")

    def run_selenium_login(self, user_id, user_pw):
        try:
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
        except Exception as e:
            raise Exception(f"로그인 실행 중 오류 발생: {str(e)}")


    def run_selenium_collection(self, stop_page):
        
        #결과화면 초기화
        self.text_area.setText("")

        current_page = 1
        results = []

        while current_page <= stop_page:
            i=1
            #페이지 정적 이동
            current_page_url = f"https://buy.11st.co.kr/my11st/order/OrderList.tmall?currpageNo={current_page}&pageNumber={current_page}" 
            self.driver.get(current_page_url)

            try:
                while i <= 10:

                    post_no = None
                    post_company = None
                    
                    if i == 1:
                        time.sleep(2)

                        tr_first = self.driver.find_element(By.CSS_SELECTOR, "#mytmall_tbl1_area > form > table > tbody > tr.first")


                        "#mytmall_tbl1_area > form > table > tbody > tr.first > td.td-center"
                        td_center = tr_first.find_element(By.CSS_SELECTOR, "td.td-center").find_element(By.XPATH, "./*[1]")
                        print("td_center 찾기완료")
                        order_state = td_center.find_element(By.XPATH, "./*[1]")
                        order_state = order_state.text.strip()
                        print("1")
                        print(order_state)
                        if order_state == "배송완료" or order_state == "구매확정"or order_state == "배송중":
                            if order_state == "구매확정":
                                post_info = td_center.find_element(By.XPATH, "./*[3]/*[1]")
                            
                            else:
                                post_info = td_center.find_element(By.XPATH, "./*[4]/*[1]")

                            try:
                                self.driver.execute_script("arguments[0].click();", post_info)
                                self.driver.switch_to.window(self.driver.window_handles[2])
                            except:
                                print("클릭에러")

                            try:
                                delivery_info = self.driver.find_element(By.CSS_SELECTOR, "#popLayWrap > div.popbody_con.tracking_delivery > div.inner.delivery > div > div > div.delivery_info > dl")
                                post_company_element = delivery_info.find_element(By.XPATH, "./*[2]/*[2]")
                                post_no_element = delivery_info.find_element(By.XPATH, "./*[3]/*[2]")
                                print("택배사찾기성공")
                                post_company = post_company_element.text.strip()
                                post_company = post_company.split()[0]
                                post_no = post_no_element.text.strip()
                                print(post_company)
                                print(post_no)
                                self.driver.close()
                                tabs = self.driver.window_handles
                                self.driver.switch_to.window(tabs[1])

                            except:
                                print("택배정보 찾기 에러")
                            

                        #상세정보 페이지 이동
                        try:
                            cust_page_selector = "#mytmall_tbl1_area > form > table > tbody > tr.first > td.first > p:nth-child(4) > a"
                            cust_page = tr_first.find_element(By.XPATH, "./*[1]/*[4]/*[1]")
                            self.driver.execute_script("arguments[0].click();", cust_page)
                            #페이지 포커스
                            tabs = self.driver.window_handles
                            self.driver.switch_to.window(tabs[1])
                        except:
                            input(f"상세정보클릭에러")

                        #고객 정보 수집
                        try:
                            cust_info_selector = "#layBody > div > div.mytmall_contArea > div.mytmall_tbl2 > table > tbody > tr > td > div"
                            cust_info_element = self.driver.find_element(By.CSS_SELECTOR, cust_info_selector)
                            cust_info = cust_info_element.text.strip()
                            cust_lines = cust_info.split("\n")
                            cust_name = cust_lines[0]
                            cust_address = cust_lines[1]
                            cust_num = cust_lines[2].strip(" /")
                            print(cust_name)
                            print(cust_address)
                            #주문관리 페이지로 이동
                            self.driver.get(current_page_url)   
                            time.sleep(1)
                        except:
                            print(f"고객정보 찾기 에러")

                    else:
                        try:
                            time.sleep(2)
                            child_no = i + 45
                            tr = self.driver.find_element(By.CSS_SELECTOR, f"#mytmall_tbl1_area > form > table > tbody > tr:nth-child({child_no})")
                            print(f"다음 상품으로 {i}")
                        except:
                            print(f"상품이동 에러 {i}")
                            
                        try:
                            td_center = tr.find_element(By.CSS_SELECTOR, "td.td-center").find_element(By.XPATH, "./*[1]")

                            order_state = td_center.find_element(By.XPATH, "./*[1]")
                            order_state = order_state.text.strip()
                            print("1")
                            print(order_state)
                        except:
                            input(f"배송조회 요소 에러 {i}")
                        
                        if order_state == "구매확정" or order_state == "배송완료" or order_state == "배송중":
                            if order_state == "구매확정":
                                post_info = td_center.find_element(By.XPATH, "./*[3]/*[1]")
                            
                            else:
                                post_info = td_center.find_element(By.XPATH, "./*[4]/*[1]")
                    

                            try:
                                self.driver.execute_script("arguments[0].click();", post_info)
                                self.driver.switch_to.window(self.driver.window_handles[2])
                            except:
                                input(f"배송조회 클릭 에러 {i}")

                            try:
                                delivery_info = self.driver.find_element(By.CSS_SELECTOR, "#popLayWrap > div.popbody_con.tracking_delivery > div.inner.delivery > div > div > div.delivery_info > dl")
                                post_company_element = delivery_info.find_element(By.XPATH, "./*[2]/*[2]")
                                post_no_element = delivery_info.find_element(By.XPATH, "./*[3]/*[2]")
                                print("택배사찾기성공")
                                post_company = post_company_element.text.strip()
                                post_company = post_company.split()[0]
                                post_no = post_no_element.text.strip()
                                print(post_company)
                                print(post_no)
                                try:
                                    self.driver.close()
                                    tabs = self.driver.window_handles
                                    self.driver.switch_to.window(tabs[1])
                                except:
                                    print(f"팝업 닫기 에러 {i}")
                            except:
                                print("배송정보 찾기 에러")

                    

                        #상세정보 페이지 이동
                        try:
                            time.sleep(1)
                            cust_page_selector = "#mytmall_tbl1_area > form > table > tbody > tr.first > td.first > p:nth-child(4) > a"
                            cust_page = tr.find_element(By.XPATH, "./*[1]/*[4]/*[1]")
                            self.driver.execute_script("arguments[0].click();", cust_page)
                            #페이지 포커스
                            tabs = self.driver.window_handles
                            self.driver.switch_to.window(tabs[1])
                        except:
                            input(f"상세정보클릭에러")

                        #고객 정보 수집
                        try:
                            cust_info_selector = "#layBody > div > div.mytmall_contArea > div.mytmall_tbl2 > table > tbody > tr > td > div"
                            cust_info_element = self.driver.find_element(By.CSS_SELECTOR, cust_info_selector)
                            cust_info = cust_info_element.text.strip()
                            cust_lines = cust_info.split("\n")
                            cust_name = cust_lines[0]
                            cust_address = cust_lines[1]
                            cust_num = cust_lines[2].strip(" /")
                            print(cust_name)
                            print(cust_address)
                            #주문관리 페이지로 이동
                            self.driver.get(current_page_url)
                            time.sleep(1)   
                        except:
                            input(f"고객정보 찾기 에러")

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
                    print("\n")
                current_page += 1    
            except:
                print("오류, 리로딩") 

        self.results = results
        self.driver.quit()
    
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
    
