import time
import logging
import pyautogui
import threading
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog
import undetected_chromedriver as uc
from undetected_chromedriver import Chrome, ChromeOptions
from tkinter import filedialog, messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

class Takedown(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.file_path = None
        self.data_frame = None
        self.progress_var = tk.DoubleVar()
        self.selected_service = tk.StringVar(value="배달의민족")
        
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
        
        # 엑셀 파일 업로드 버튼 생성
        self.upload_button = ttk.Button(self, text="엑셀 파일 업로드", command=self.upload_file, style='TButton')
        self.upload_button.pack(pady=10)
        
        # 프레임 생성
        self.service_frame = ttk.Frame(self, style='TFrame')
        self.service_frame.pack(pady=10)
        
        self.baemin_radio = ttk.Radiobutton(self.service_frame, text="배달의민족", variable=self.selected_service, value="배달의민족", style='TRadiobutton')
        self.baemin_radio.pack(side=tk.LEFT)
        
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
        
        # 진행률 ( % ) 확인 용
        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, maximum=100, style='TProgressbar.Horizontal.TProgressbar')
        self.progress_bar.pack(pady=10, padx=10, fill=tk.X)
        
        self.progress_label = ttk.Label(self, text="진행 상황: 0%", style='TLabel')
        self.progress_label.pack(pady=10)
        
        self.treeview_frame = ttk.Frame(self, style='TFrame')
        self.treeview_frame.pack(fill="both", expand=True)
        
        self.treeview = ttk.Treeview(self.treeview_frame, style='Treeview')
        self.treeview.pack(fill="both", expand=True)

    #화면상의 엑셀 상의 데이터를 삭제하는 작업 진행 
    def clear_treeview(self):
        self.treeview.delete(*self.treeview.get_children())

    #엑셀 업로드 기능 
    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        if self.file_path:
            self.data_frame = pd.read_excel(self.file_path)
            service = self.selected_service.get()
            expected_columns = []
            if service == "배달의민족":
                expected_columns = ["SHOP_ID", "Review_Number"]
            elif service == "쿠팡이츠":
                expected_columns = ["SHOP_ID", "Business Number", "Order_Number", "Reason"]

            if all(column in self.data_frame.columns for column in expected_columns):
                self.display_excel()
            else:
                messagebox.showwarning("경고", f"선택된 서비스({service})의 엑셀 형식이 올바르지 않습니다.")
                self.data_frame = None

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
        
        driver = Chrome(options=options, use_subprocess=True)
        try:
            baemin_url = "https://design.happytalkio.com/chatting?siteId=4000000024&siteName=%EC%9A%B0%EC%95%84%ED%95%9C%ED%98%95%EC%A0%9C%EB%93%A4&categoryId=61602&divisionId=200880&partnerId=&shopId=&params="
            driver.get(baemin_url)
            driver.delete_all_cookies()
            time.sleep(3)
            
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
        
        driver = Chrome(options=options, use_subprocess=True)
        try:
            coupang_url = "https://design.happytalkio.com/chatting?siteId=4000002553&siteName=%EC%BF%A0%ED%8C%A1%EC%9D%B4%EC%B8%A0&categoryId=154858&divisionId=155774&partnerId=&shopId=&params="
            driver.get(coupang_url)
            driver.delete_all_cookies()
            time.sleep(3)
            
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
            self.data_frame.to_excel(self.file_path, index=False)
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
        
        driver = Chrome(options=options, use_subprocess=True)
        try:
            # 크롤링 작업 선택 ( 배달의민족 / 쿠팡이츠 / 땡겨요 )
            if selected_service == "배달의민족":
                self.takedown_baemin(driver, row, row_idx)
            elif selected_service == "쿠팡이츠":
                self.takedown_coupang(driver, row, row_idx)
        except Exception as e:
            print(f"에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"에러 발생: {e}"
        finally:
            driver.quit()

    def takedown_baemin(self, driver, row, row_idx):
        try:
            baemin_class = self.baemin_label.cget("text")
            if not baemin_class:
                raise Exception("배달의민족 클래스 이름이 설정되지 않았습니다.")
            
            # 배달의민족 리뷰게시중단 사이트로 이동 
            driver.get("https://design.happytalkio.com/chatting?siteId=4000000024&siteName=%EC%9A%B0%EC%95%84%ED%95%9C%ED%98%95%EC%A0%9C%EB%93%A4&categoryId=61602&divisionId=200880&partnerId=&shopId=&params=")
            time.sleep(3)
            
            # ► 리뷰 게시중단 신청 버튼 클릭 
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='► 리뷰 게시중단 신청']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 시작하기 버튼 클릭 
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='► 시작하기']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 확인했어요 클릭 
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='확인했어요.']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # SHOP_ID 입력
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[2]/div/textarea"))
            )
            textarea.send_keys(str(row["SHOP_ID"]))
            time.sleep(2)   
            
            # 입력 버튼 클릭
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]'))
            )
            clip_button.click()
            time.sleep(2)
            
            # Review_Number 입력
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div/div/div[2]/div/textarea"))
            )
            textarea.send_keys(str(row["Review_Number"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div/div[3]'))
            )
            clip_button.click()
            time.sleep(2)
            
            # 대표자 클릭 
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='대표자']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 이메일 클릭 
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='이메일']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            # 접수하기 클릭 
            baemin_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{baemin_class}')]/button[text()='► 접수하기']"))
            )
            baemin_button.click()
            time.sleep(2)
            
            self.data_frame.at[row_idx, '상태'] = "성공"

        except Exception as e:
            logging.error(f"배달의민족 에러 발생: {e}")
            self.data_frame.at[row_idx, '상태'] = f"배달의민족 에러 발생: {e}"

    # 쿠팡이츠 리뷰 게시중단 신청 작업 
    def takedown_coupang(self, driver, row, row_idx):
        try:
            coupang_class = self.coupang_label.cget("text")
            if not coupang_class:
                raise Exception("쿠팡이츠 클래스 이름이 설정되지 않았습니다.")
            
            # 쿠팡이츠 리뷰게시중단 사이트로 이동 
            driver.get("https://design.happytalkio.com/chatting?siteId=4000002553&siteName=%EC%BF%A0%ED%8C%A1%EC%9D%B4%EC%B8%A0&categoryId=154858&divisionId=155774&partnerId=&shopId=&params=")
            time.sleep(3)
            
            # 리뷰 블라인드/게시중단 요청 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='리뷰 블라인드/게시중단 요청']"))
            )
            coupang_button.click()
            time.sleep(2)
            
            # 게시중단 요청만 신청 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='게시중단 요청만 신청']"))
            )
            coupang_button.click()
            time.sleep(2)

            # 본인신청 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='본인 신청']"))
            )
            coupang_button.click()
            time.sleep(2)
            
            # ► 간편하게 접수하기 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='▶ 간편하게 접수하기']"))
            )
            coupang_button.click()
            time.sleep(2)

            # ► 계속 신청하기 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='▶ 계속 신청하기']"))
            )
            coupang_button.click()
            time.sleep(2)

            # SHOP_ID 입력
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["SHOP_ID"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # Business Number 입력
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["Business Number"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # 일치하며 이어서 진행하기 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='일치하며 이어서 진행하기']"))
            )
            coupang_button.click()
            time.sleep(2)
            
            # Order_Number 입력
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["Order_Number"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # 기타 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='기타']"))
            )
            coupang_button.click()
            time.sleep(2)

            # Reason 입력
            textarea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[3]/div/textarea"))
            )
            textarea.send_keys(str(row["Reason"]))   
            time.sleep(2)
            
            # 입력 버튼 클릭
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[4]'))
            )
            clip_button.click()
            time.sleep(2)

            # 네 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='네']"))
            )
            coupang_button.click()
            time.sleep(2)

            # ▶ 동의하고 접수하기 버튼 클릭 
            coupang_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, '{coupang_class}')]/button[text()='▶ 동의하고 접수하기']"))
            )
            coupang_button.click()
            time.sleep(2)

            self.data_frame.at[row_idx, '상태'] = "성공"
            
        except Exception as e:
            print(f"쿠팡이츠 에러 발생: {e}")
            self.data_frame.at[row_idx, '상태'] = f"쿠팡이츠 에러 발생: {e}"

# Initialize UI
if __name__ == '__main__':
    root = tk.Tk()
    root.title("크롤링 프로그램")
    root.geometry("1000x600")
    frame = Takedown(root)
    frame.pack(fill="both", expand=True)
    root.mainloop()