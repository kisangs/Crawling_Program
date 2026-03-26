import time
import logging
import pyautogui
import threading
import pandas as pd
import tkinter as tk
import requests
from io import BytesIO
from tkinter import ttk, filedialog
import undetected_chromedriver as uc
from undetected_chromedriver import Chrome, ChromeOptions
from tkinter import filedialog, messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    WebDriverException
)

class Takedown(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.file_path = None
        self.data_frame = None
        self.progress_var = tk.DoubleVar()
        self.selected_service = tk.StringVar(value="배달의민족")

        self.google_xlsx_url = "https://docs.google.com/spreadsheets/d/1y8Wj6B_i3H2EhZnWGT66k76rgxnRppvhZHzLPo76jdE/export?format=xlsx"
        self.apps_script_url = "https://script.google.com/a/macros/cloudkitchens.com/s/AKfycbx5w8EJdqABT6VKyZ3pHVj5VwLtJux6xAFwBNxncZ8kmXkZtOs-Yfe40c0WNzLjyOYz/exec"

        self.baemin_sheet_name = "배달의민족"
        self.coupang_sheet_name = "쿠팡이츠"
        self.baemin_result_sheet_name = "배달의민족_결과"
        self.coupang_result_sheet_name = "쿠팡이츠_결과"
        
        # 다크 모드 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        
        # 다크 모드 스타일 설정
        style.configure('TFrame', background='#2b2b2b')
        style.configure('TButton', background='#ffffff', foreground='#000000', font=('Helvetica', 10, 'bold'))
        style.configure('TLabel', background='#2b2b2b', foreground='#ffffff')
        style.configure('TProgressbar', foreground='#00ff00', background='#3c3c3c')
        style.configure('TProgressbar.Horizontal.TProgressbar', troughcolor='#3c3c3c', background='#00ff00')
        style.configure('Treeview', background='#2b2b2b', foreground='#ffffff', fieldbackground='#2b2b2b', bordercolor='#ffffff')
        style.map('Treeview', background=[('selected', '#3c3c3c')], foreground=[('selected', '#ffffff')])
        style.configure('Treeview.Heading', background='#3c3c3c', foreground='#ffffff')
        style.configure('TEntry', fieldbackground='#ffffff', background='#2b2b2b', foreground='#000000')
        
        # 엑셀 파일 업로드 버튼 ( 배달의민족 / 쿠팡이츠 ) 생성
        self.upload_frame = ttk.Frame(self, style='TFrame')
        self.upload_frame.pack(pady=10)

        #배달의민족 엑셀 파일 불러오기 버튼 
        self.baemin_upload_button = ttk.Button(
            self.upload_frame,
            text="배달의민족 불러오기",
            command=self.load_baemin_file,
            style='TButton'
        )
        self.baemin_upload_button.pack(side=tk.LEFT, padx=5)

        #쿠팡이츠 엑셀 파일 불러오기 버튼
        self.coupang_upload_button = ttk.Button(
            self.upload_frame,
            text="쿠팡이츠 불러오기",
            command=self.load_coupang_file,
            style='TButton'
        )
        self.coupang_upload_button.pack(side=tk.LEFT, padx=5)
        
        # 프레임 생성
        self.service_frame = ttk.Frame(self, style='TFrame')
        self.service_frame.pack(pady=10)
        
        #배달의민족 라디오 버튼 생성
        self.baemin_radio = ttk.Radiobutton(self.service_frame, text="배달의민족", variable=self.selected_service, value="배달의민족", style='TRadiobutton')
        self.baemin_radio.pack(side=tk.LEFT)

        #쿠팡이츠 라디오 버튼 생성
        self.coupang_radio = ttk.Radiobutton(self.service_frame, text="쿠팡이츠", variable=self.selected_service, value="쿠팡이츠", style='TRadiobutton')
        self.coupang_radio.pack(side=tk.LEFT)

        # label 칸 / 검색 버튼들 프레임 생성
        self.input_frame = ttk.Frame(self, style='TFrame')
        self.input_frame.pack(pady=0, padx=0)
        
        # label 칸 / 검색 버튼들 프레임 생성
        self.input_frame = ttk.Frame(self, style='TFrame')
        self.input_frame.pack(pady=0, padx=0)

        # 배달의민족 Class 값을 받아올 label 생성
        self.baemin_label = ttk.Label(self.input_frame, text="배달의민족 Class 확인 필요", style='TLabel', background='#2b2b2b', foreground='#ffffff', anchor='w')
        self.baemin_label.pack(side=tk.LEFT, pady=0, padx=10, fill=tk.X, expand=True)

        #배달의민족 Class 검색 버튼 
        self.baemin_search_button = ttk.Button(self.input_frame, text="검색", command=self.Baemin_class_Search, style='TButton')
        self.baemin_search_button.pack(side=tk.LEFT, pady=0, padx=10)

        # 쿠팡이츠 Class 값을 받아올 label 생성
        self.coupang_label = ttk.Label(self.input_frame, text="쿠팡이츠 Class 확인 필요", style='TLabel', background='#2b2b2b', foreground='#ffffff', anchor='w')
        self.coupang_label.pack(side=tk.LEFT, pady=0, padx=10, fill=tk.X, expand=True)

        #쿠팡이츠 Class 검색 버튼
        self.coupang_search_button = ttk.Button(self.input_frame, text="검색", command=self.Coupang_class_Search, style='TButton')
        self.coupang_search_button.pack(side=tk.LEFT, pady=0, padx=0)
        self.start_button = ttk.Button(self, text="시작", command=self.start_thread, style='TButton')
        self.start_button.pack(pady=10)

        #결과 재전송 버튼 
        self.retry_upload_button = ttk.Button(self,text="결과 재전송",command=self.retry_upload_results,style='TButton')
        self.retry_upload_button.pack(pady=5)
        
        # 진행률 ( % ) 확인 용
        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, maximum=100, style='TProgressbar.Horizontal.TProgressbar')
        self.progress_bar.pack(pady=10, padx=10, fill=tk.X)
        
        self.progress_label = ttk.Label(self, text="진행 상황: 0%", style='TLabel')
        self.progress_label.pack(pady=10)
        
        self.treeview_frame = ttk.Frame(self, style='TFrame')
        self.treeview_frame.pack(fill="both", expand=True)
        
        self.treeview = ttk.Treeview(self.treeview_frame, style='Treeview')
        self.treeview.pack(fill="both", expand=True)

    #오류 원인 분석 
    def get_user_friendly_error(self, step, exception):
        if isinstance(exception, TimeoutException):
            if "클릭" in step:
                return f"{step} 단계에서 버튼이 나타나지 않거나 클릭 가능한 상태가 되지 않았습니다"
            elif "입력" in step:
                return f"{step} 단계에서 입력창이 나타나지 않았습니다"
            else:
                return f"{step} 단계에서 페이지 또는 요소 응답 시간이 초과되었습니다"

        elif isinstance(exception, NoSuchElementException):
            return f"{step} 단계에서 필요한 화면 요소를 찾지 못했습니다"

        elif isinstance(exception, ElementClickInterceptedException):
            return f"{step} 단계에서 다른 요소에 가려 클릭할 수 없었습니다"

        elif isinstance(exception, ElementNotInteractableException):
            return f"{step} 단계에서 요소가 비활성화 상태이거나 입력/클릭할 수 없었습니다"

        elif isinstance(exception, StaleElementReferenceException):
            return f"{step} 단계에서 화면이 갱신되어 요소를 다시 찾아야 했습니다"

        elif isinstance(exception, WebDriverException):
            return f"{step} 단계에서 브라우저 동작 중 오류가 발생했습니다"

        else:
            return f"{step} 단계에서 알 수 없는 오류가 발생했습니다"
            
    #화면상의 엑셀 상의 데이터를 삭제하는 작업 진행 
    def clear_treeview(self):
        self.treeview.delete(*self.treeview.get_children())

    #엑셀 업로드 기능 
    def load_baemin_file(self):
        self.selected_service.set("배달의민족")
        self.load_sheet_from_google(self.baemin_sheet_name)

    def load_coupang_file(self):
        self.selected_service.set("쿠팡이츠")
        self.load_sheet_from_google(self.coupang_sheet_name)

    def load_sheet_from_google(self, sheet_name):
        try:
            response = requests.get(self.google_xlsx_url, timeout=30)
            response.raise_for_status()

            excel_data = BytesIO(response.content)
            self.data_frame = pd.read_excel(excel_data, sheet_name=sheet_name)
            self.data_frame.columns = self.data_frame.columns.astype(str).str.strip()

            if '상태' not in self.data_frame.columns:
                self.data_frame['상태'] = ''
            if '에러' not in self.data_frame.columns:
                self.data_frame['에러'] = ''

            service = self.selected_service.get()
            expected_columns = []

            if service == "배달의민족":
                expected_columns = ["SHOP_ID", "Review_Number"]
            elif service == "쿠팡이츠":
                expected_columns = ["SHOP_ID", "Business Number", "Order_Number", "Reason"]

            if all(column in self.data_frame.columns for column in expected_columns):
                self.display_excel()
                messagebox.showinfo("완료", f"{sheet_name} 탭 데이터를 불러왔습니다.")
            else:
                messagebox.showwarning("경고", f"선택된 서비스({service})의 시트 형식이 올바르지 않습니다.")
                self.data_frame = None

        except Exception as e:
            messagebox.showerror("오류", f"구글 시트 데이터 불러오기 실패:\n{e}")
            self.data_frame = None

    #결과 재전송 로직 
    def retry_upload_results(self):
        if self.data_frame is None:
            messagebox.showwarning("경고", "재전송할 데이터가 없습니다.")
            return

        confirm = messagebox.askyesno("확인", "현재 화면의 결과 데이터를 다시 전송하시겠습니까?")
        if not confirm:
            return

        self.upload_results_to_apps_script()

    #화면 상에 엑셀을 표시 (Treeview)
    def display_excel(self):
        if self.data_frame is not None:
            self.clear_treeview()
            self.treeview["column"] = list(self.data_frame.columns)
            self.treeview["show"] = "headings"
            for column in self.treeview["columns"]:
                self.treeview.heading(column, text=column)
            for row in self.data_frame.to_numpy().tolist():
                self.treeview.insert("", "end", values=row)

    #진행중인 행을 표시할 수 있게 하이라이트 표시 
    def highlight_row(self, row_id):
        """주어진 행 ID를 하이라이트"""
        for item in self.treeview.get_children():
            self.treeview.item(item, tags="")
        self.treeview.item(row_id, tags=("highlight",))
        self.treeview.tag_configure("highlight", background="yellow", foreground="black")

    #Apps Script 업로드
    def upload_results_to_apps_script(self):
        try:
            if self.data_frame is None:
                messagebox.showwarning("경고", "업로드할 데이터가 없습니다.")
                return

            service = self.selected_service.get()

            if service == "배달의민족":
                result_sheet_name = self.baemin_result_sheet_name
            elif service == "쿠팡이츠":
                result_sheet_name = self.coupang_result_sheet_name
            else:
                raise Exception("알 수 없는 서비스입니다.")

            upload_df = self.data_frame.copy()
            upload_df = upload_df.fillna("")

            payload = {
                "sheetName": result_sheet_name,
                "columns": list(upload_df.columns),
                "rows": upload_df.astype(str).values.tolist()
            }

            response = requests.post(self.apps_script_url, json=payload, timeout=30)
            response.raise_for_status()

            result = response.json()

            if result.get("success"):
                messagebox.showinfo("업로드 완료", f"{result_sheet_name} 탭에 결과 업로드 완료")
            else:
                messagebox.showerror("업로드 실패", result.get("message", "알 수 없는 오류"))

        except Exception as e:
            messagebox.showerror("업로드 오류", f"Apps Script 업로드 실패:\n{e}")

    #배달의민족 리뷰 게시중단 사이트의 Button Class 받아오는 작업 
    def Baemin_class_Search(self):
        #Chrome Driver 정의 
        options = ChromeOptions()
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--disable-plugins-discovery")
        options.add_argument("--incognito")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--log-level=3')
        
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        try:
            baemin_url = "https://design.happytalkio.com/chatting?siteId=4000000024&siteName=%EC%9A%B0%EC%95%84%ED%95%9C%ED%98%95%EC%A0%9C%EB%93%A4&categoryId=61602&divisionId=200880&partnerId=&shopId=&params="
            driver.get(baemin_url)
            driver.delete_all_cookies()
            time.sleep(3)
            
            WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//li[@data-idx="1"]/button')))
            elements = driver.find_elements(By.XPATH, '//li[@data-idx="1"]/button')
            if elements:
                button_element = elements[0]
                parent_element = button_element.find_element(By.XPATH, "..")
                class_name = parent_element.get_attribute("class")
                if class_name:
                    self.baemin_label.config(text=class_name)
                else:
                    logging.error("클래스가 없음")
            else:
                logging.error("Label이 없음")
                
        except Exception as e:
            logging.error(f"클래스 이름을 찾는 중 별도 오류 발생: {e}")
        finally:
            driver.quit()

    #쿠팡이츠 리뷰 게시중단 사이트의 Button Class 받아오는 작업 
    def Coupang_class_Search(self):
        options = ChromeOptions()
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--disable-plugins-discovery")
        options.add_argument("--incognito")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--log-level=3')
        
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        try:
            coupang_url = "https://design.happytalkio.com/chatting?siteId=4000002553&siteName=%EC%BF%A0%ED%8C%A1%EC%9D%B4%EC%B8%A0&categoryId=154858&divisionId=155774&partnerId=&shopId=&params="
            driver.get(coupang_url)
            driver.delete_all_cookies()
            time.sleep(3)
            
            WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//li[@data-idx="1"]/button')))
            elements = driver.find_elements(By.XPATH, '//li[@data-idx="1"]/button')
            if elements:
                button_element = elements[0]
                parent_element = button_element.find_element(By.XPATH, "..")
                class_name = parent_element.get_attribute("class")
                if class_name:
                    self.coupang_label.config(text=class_name)
                else:
                    logging.error("클래스가 없음")
            else:
                logging.error("Label이 없음")
                
        except Exception as e:
            logging.error(f"클래스 이름을 찾는 중 별도 오류 발생: {e}")
        finally:
            driver.quit()

    def start_thread(self):
        threading.Thread(target=self.start_crawling).start()

    def start_crawling(self):
        if self.data_frame is not None:
            service = self.selected_service.get()
            if (service == "배달의민족" and self.baemin_label.cget("text") == "배달의민족 Class 확인 필요") or \
            (service == "쿠팡이츠" and self.coupang_label.cget("text") == "쿠팡이츠 Class 확인 필요"):
                messagebox.showwarning("경고", f"{service}의 Class 항목이 설정되지 않았습니다.")
                return

            total_rows = len(self.data_frame.index)
            for index, row in self.data_frame.iterrows():
                item_id = self.treeview.get_children()[index]  # 해당 행의 ID
                self.highlight_row(item_id)  # 현재 행을 하이라이트

                self.crawl_site(row, index)

                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"진행 상황: {progress:.2f}%")
                self.update_idletasks()

            # 데이터 저장 후 엑셀에 저장
            self.upload_results_to_apps_script()
            self.display_excel()
            messagebox.showinfo("완료", "완료됨~")
        else:
            messagebox.showwarning("경고", "엑셀파일부터 업로드 하셔야죠~")

    def scroll_to_element(self, driver, element):
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(1)

    def crawl_site(self, row, row_idx):
        selected_service = self.selected_service.get()
        options = ChromeOptions()
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--disable-plugins-discovery")
        options.add_argument("--incognito")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--log-level=3')
        
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        try:
            # 크롤링 작업 선택 ( 배달의민족 / 쿠팡이츠 / 땡겨요 )
            if selected_service == "배달의민족":
                self.takedown_baemin(driver, row, row_idx)
            elif selected_service == "쿠팡이츠":
                self.takedown_coupang(driver, row, row_idx)
        except Exception as e:
            step = "크롤링 실행"
            friendly_error = self.get_user_friendly_error(step, e)
            logging.error(f"[{step}] {type(e).__name__}: {e}")
            self.data_frame.at[row_idx, '상태'] = "실패"
            self.data_frame.at[row_idx, '에러'] = friendly_error
        finally:
            driver.quit()

    def takedown_baemin(self, driver, row, row_idx):
        step = "초기화"
        try:
            step = "배달의민족 Class 확인"
            baemin_class = self.baemin_label.cget("text")
            if not baemin_class:
                raise Exception("배달의민족 클래스 이름이 설정되지 않았습니다.")
            
            # 배달의민족 리뷰게시중단 사이트로 이동 
            step = "배달의민족 사이트 접속"
            driver.get("https://design.happytalkio.com/chatting?siteId=4000000024&siteName=%EC%9A%B0%EC%95%84%ED%95%9C%ED%98%95%EC%A0%9C%EB%93%A4&categoryId=61602&divisionId=200880&partnerId=&shopId=&params=")
            time.sleep(3)

            #초기 버튼 로딩 확인 대기 
            step = "초기 버튼 로딩 확인"
            WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//li[@data-idx="1"]/button')))
            time.sleep(2)
            
            # 리뷰게시중단/리뷰케어 신청 버튼 클릭 
            step = "리뷰게시중단/리뷰케어 신청 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='리뷰게시중단/리뷰케어 신청']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 리뷰게시중단 신청 버튼 클릭 
            step = "리뷰게시중단 신청 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='리뷰게시중단 신청']"))
            )
            baemin_button.click()
            time.sleep(2)

            # 시작하기 버튼 클릭 
            step = "시작하기 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='► 시작하기']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 확인했어요 클릭 
            step = "확인했어요 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='확인했어요.']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # SHOP_ID 입력
            step = "SHOP_ID 입력"
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[2]/div/textarea"))
            )
            textarea.send_keys(str(row["SHOP_ID"]))
            time.sleep(2)   
            
            # 입력 버튼 클릭
            step = "SHOP_ID 입력 버튼 클릭"
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]'))
            )
            clip_button.click()
            time.sleep(2)
            
            # Review_Number 입력
            step = "Review_Number 입력"
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[2]/div/textarea"))
            )
            textarea.send_keys(str(row["Review_Number"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            step = "Review_Number 입력 버튼 클릭"
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]'))
            )
            clip_button.click()
            time.sleep(2)
            
            # 대표자 클릭 
            step = "대표자 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='대표자']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 이메일 클릭 
            step = "이메일 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='이메일']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 접수하기 클릭
            step = "접수하기 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='► 접수하기']"))
            )
            baemin_button.click()
            time.sleep(2)

            # 종료하기 클릭
            step = "종료하기 버튼 클릭"
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='► 종료하기']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            self.data_frame.at[row_idx, '상태'] = "성공"
            self.data_frame.at[row_idx, '에러'] = ""

        except Exception as e:
            friendly_error = self.get_user_friendly_error(step, e)
            logging.error(f"[배달의민족][{step}] {type(e).__name__}: {e}")
            self.data_frame.at[row_idx, '상태'] = "실패"
            self.data_frame.at[row_idx, '에러'] = friendly_error

    # 쿠팡이츠 리뷰 게시중단 신청 작업 
    def takedown_coupang(self, driver, row, row_idx):
        step = "초기화"
        try:
            step = "쿠팡이츠 클래스 확인"
            coupang_class = self.coupang_label.cget("text")
            if not coupang_class:
                raise Exception("쿠팡이츠 클래스 이름이 설정되지 않았습니다.")
            
            # 쿠팡이츠 리뷰게시중단 사이트로 이동 
            step = "쿠팡이츠 페이지 접속"
            driver.get("https://design.happytalkio.com/chatting?siteId=4000002553&siteName=%EC%BF%A0%ED%8C%A1%EC%9D%B4%EC%B8%A0&categoryId=154858&divisionId=155774&partnerId=&shopId=&params=")
            time.sleep(3)

            #초기 버튼 로딩 확인 대기 
            step = "초기 버튼 로딩 확인"
            WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//li[@data-idx="1"]/button')))
            
            # 리뷰 블라인드/게시중단 요청 버튼 클릭 
            step = "리뷰 블라인드/게시중단 요청 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='리뷰 블라인드/게시중단 요청']"))
            )
            coupang_button.click()
            time.sleep(2)
            
            # 게시중단 요청만 신청 버튼 클릭 
            step = "게시중단 요청만 신청 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='게시중단 요청만 신청']"))
            )
            coupang_button.click()
            time.sleep(2)

            # 본인신청 버튼 클릭 
            step = "본인 신청 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='본인 신청']"))
            )
            coupang_button.click()
            time.sleep(2)
            
            # ► 간편하게 접수하기 버튼 클릭 
            step = "간편하게 접수하기 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='▶ 간편하게 접수하기']"))
            )
            coupang_button.click()
            time.sleep(2)

            # ► 계속 신청하기 버튼 클릭 
            step = "계속 신청하기 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='▶ 계속 신청하기']"))
            )
            coupang_button.click()
            time.sleep(2)

            # SHOP_ID 입력
            step = "SHOP_ID 입력"
            textarea = WebDriverWait(driver, 10).until(    
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["SHOP_ID"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            step = "SHOP_ID 입력 버튼 클릭"
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # Business Number 입력
            step = "Business Number 입력"
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["Business Number"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            step = "Business Number 입력 버튼 클릭"
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # 일치하며 이어서 진행하기 버튼 클릭 
            step = "일치하며 이어서 진행하기 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='일치하며 이어서 진행하기']"))
            )
            coupang_button.click()
            time.sleep(2)
            
            # Order_Number 입력
            step = "Order_Number 입력"
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["Order_Number"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            step = "Order_Number 입력 버튼 클릭"
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # 기타 버튼 클릭 
            step = "기타 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='기타']"))
            )
            coupang_button.click()
            time.sleep(2)

            # Reason 입력
            step = "Reason 입력"
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["Reason"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            step = "Reason 입력 버튼 클릭"
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # 네 버튼 클릭 
            step = "네 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='네']"))
            )
            coupang_button.click()
            time.sleep(2)

            # ▶ 동의하고 접수하기 버튼 클릭 
            step = "동의하고 접수하기 버튼 클릭"
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='▶ 동의하고 접수하기']"))
            )
            coupang_button.click()
            time.sleep(2)

            self.data_frame.at[row_idx, '상태'] = "성공"
            self.data_frame.at[row_idx, '에러'] = ""
            
        except Exception as e:
            friendly_error = self.get_user_friendly_error(step, e)
            logging.error(f"[쿠팡이츠][{step}] {type(e).__name__}: {e}")
            self.data_frame.at[row_idx, '상태'] = "실패"
            self.data_frame.at[row_idx, '에러'] = friendly_error

# Initialize UI
if __name__ == '__main__':
    root = tk.Tk()
    root.title("크롤링 프로그램")
    root.geometry("1000x600")
    frame = Takedown(root)
    frame.pack(fill="both", expand=True)
    root.mainloop()