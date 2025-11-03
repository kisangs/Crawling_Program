import time
import threading
import pandas as pd
import tkinter as tk

from selenium import webdriver
from tkinter import ttk, filedialog
from tkinter import filedialog, messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from undetected_chromedriver import Chrome, ChromeOptions
from selenium.webdriver.support import expected_conditions as EC


class Crawling(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.file_path = None
        self.data_frame = None
        self.progress_var = tk.DoubleVar()
        self.selected_service = tk.StringVar(value="배달의민족")
        # 다크 모드 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        # 전체 화면 어두운 색
        style.configure('TFrame', background='#2b2b2b')
        
        # 버튼의 색은 하얀색, 글씨는 검은색
        style.configure('TButton', background='#ffffff', foreground='#000000', font=('Helvetica', 12, 'bold'))
        
        # 전체 글씨 하얀색
        style.configure('TLabel', background='#2b2b2b', foreground='#ffffff')
        
        # 진행률은 초록색, 글씨는 하얀색
        style.configure('TProgressbar', foreground='#00ff00', background='#3c3c3c')
        style.configure('TProgressbar.Horizontal.TProgressbar', troughcolor='#3c3c3c', background='#00ff00')
        # Treeview 스타일 설정
        style.configure('Treeview', background='#2b2b2b', foreground='#ffffff', fieldbackground='#2b2b2b', bordercolor='#ffffff')
        style.map('Treeview', background=[('selected', '#3c3c3c')], foreground=[('selected', '#ffffff')])
        style.configure('Treeview.Heading', background='#3c3c3c', foreground='#ffffff')

        self.upload_button = ttk.Button(self, text="엑셀 파일 업로드", command=self.upload_file, style='TButton')
        self.upload_button.pack(pady=10)
        
        self.service_frame = ttk.Frame(self, style='TFrame')
        self.service_frame.pack(pady=10)
        
        self.baemin_radio = ttk.Radiobutton(self.service_frame, text="배달의민족", variable=self.selected_service, value="배달의민족", style='TRadiobutton')
        self.baemin_radio.pack(side=tk.LEFT)
        
        self.coupang_radio = ttk.Radiobutton(self.service_frame, text="쿠팡이츠", variable=self.selected_service, value="쿠팡이츠", style='TRadiobutton')
        self.coupang_radio.pack(side=tk.LEFT)
        
        self.ddangyo_radio = ttk.Radiobutton(self.service_frame, text="땡겨요", variable=self.selected_service, value="땡겨요", style='TRadiobutton')
        self.ddangyo_radio.pack(side=tk.LEFT)
        
        self.start_button = ttk.Button(self, text="시작", command=self.start_thread, style='TButton')
        self.start_button.pack(pady=10)

        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, maximum=100, style='TProgressbar.Horizontal.TProgressbar')
        self.progress_bar.pack(pady=10, padx=10, fill=tk.X)
        
        self.progress_label = ttk.Label(self, text="진행 상황: 0%", style='TLabel')
        self.progress_label.pack(pady=10)
        
        self.treeview_frame = ttk.Frame(self, style='TFrame')
        self.treeview_frame.pack(fill="both", expand=True)
        
        self.treeview = ttk.Treeview(self.treeview_frame, style='Treeview')
        self.treeview.pack(fill="both", expand=True)

    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        if self.file_path:
            self.data_frame = pd.read_excel(self.file_path)
            self.display_excel()

    def display_excel(self):
        if self.data_frame is not None:
            self.treeview["column"] = list(self.data_frame.columns)
            self.treeview["show"] = "headings"
            for column in self.treeview["columns"]:
                self.treeview.heading(column, text=column)
            for row in self.data_frame.to_numpy().tolist():
                self.treeview.insert("", "end", values=row)

    def start_thread(self):
        threading.Thread(target=self.start_crawling).start()

    def start_crawling(self):
        if self.data_frame is not None:
            total_rows = len(self.data_frame.index)
            for index, row in self.data_frame.iterrows():
                self.crawl_site(row["ID"], row["PW"], index)
                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"진행 상황: {progress:.2f}%")
                self.update_idletasks()
            # 크롤링 작업이 끝나면 결과를 엑셀 파일로 저장
            self.data_frame.to_excel(self.file_path, index=False)
            messagebox.showinfo("완료", "크롤링 작업이 완료되었습니다.")
        else:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 업로드하세요.")

    def scroll_to_element(self, driver, element):
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(1)

    def crawl_site(self, user_id, user_pw, row_idx):
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

        
        service = Service(ChromeDriverManager().install())
        driver = Chrome(service=service, options=options)
        driver.implicitly_wait(10)
        driver.delete_all_cookies()

        # 배민 / 쿠팡이츠 / 땡겨요 선택 할 수 있도록 세팅 
        try:
            if selected_service == "배달의민족":
                self.crawl_baemin(driver, user_id, user_pw, row_idx)
            elif selected_service == "쿠팡이츠":
                self.crawl_coupang(driver, user_id, user_pw, row_idx)
            elif selected_service == "땡겨요":
                self.crawl_ddangyo(driver, user_id, user_pw, row_idx)
        except Exception as e:
            print(f"에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"에러 발생: {e}"
        finally:
            driver.quit()

    # 배달의민족 크롤링 작업 ( 로그인 -> 팝업 닫기 -> 연락처 수집 -> 사업자번호 수집 -> 주소 수집 -> 메인 화면으로 이동 -> 가게명 및 가게번호 수집 )
    def crawl_baemin(self, driver, user_id, user_pw, row_idx):
        try:
            driver.get("https://self.baemin.com/mypage/owner")
            time.sleep(5)
            
            # 로그인 입력
            id_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[1]/div/div/form/div[1]/span/input')))
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(5)
            
            #비밀번호 입력
            pw_input = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/form/div[2]/span/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(5)

            #로그인 버튼 클릭
            login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div/form/button')))
            login_button.click()
            time.sleep(5)

            # 로그인 실패 시 C열에 로그인 실패 기록 후 다음 작업 진행 
            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 또는 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.data_frame.at[row_idx, '연락처'] = "로그인 실패"
                driver.quit()
                return
            time.sleep(5)
            
            # "닫기" 버튼 클릭 처리, 화면에 표시된 버튼만 클릭
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                    time.sleep(5)

            # 연락처 값 크롤링
            try:
                contact_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div[1]/form/div/div/div[2]/div')
                self.scroll_to_element(driver, contact_element)
                contact_value = contact_element.text
                self.data_frame.at[row_idx, '연락처'] = contact_value
            except Exception as e:
                self.data_frame.at[row_idx, '연락처'] = f"에러 발생: {e}"
            time.sleep(5)

            # 사업자번호 값 크롤링
            try:
                business_number_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[1]/div[6]/div[5]/div')
                self.scroll_to_element(driver, business_number_element)
                business_number_value = business_number_element.text
                self.data_frame.at[row_idx, '사업자번호'] = business_number_value
            except Exception as e:
                self.data_frame.at[row_idx, '사업자번호'] = f"에러 발생: {e}"
            time.sleep(5)

            # 주소 값 크롤링
            try:
                address_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[1]/div[6]/div[12]/dl/dd[2]')
                self.scroll_to_element(driver, address_element)
                address_value = address_element.text
                self.data_frame.at[row_idx, '주소'] = address_value
            except Exception as e:
                self.data_frame.at[row_idx, '주소'] = f"에러 발생: {e}"
            time.sleep(5)

            # 배민 메인 화면으로 이동 
            driver.get("https://self.baemin.com/")
            time.sleep(5)
            
            # "닫기" 버튼 클릭 처리, 화면에 표시된 버튼만 클릭
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                    time.sleep(5)

            # 가게 정보 크롤링 
            try:
                select_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div[1]/div[4]/div[1]/div[1]/div/div/select')))
                options = select_element.find_elements(By.TAG_NAME, 'option')
                for idx, option in enumerate(options):
                    self.data_frame.at[row_idx, f'가게명_{idx+1}'] = option.text
            except Exception as e:
                self.data_frame.at[row_idx, '가게명'] = f"에러 발생: {e}"
            time.sleep(5)
            
            # 쿠키 삭제 및 종료 
            driver.delete_all_cookies()
        except Exception as e:
            print(f"배달의민족 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"배달의민족 에러 발생: {e}"

    # 쿠팡이츠 크롤링 작업 ( 로그인 -> 가게명 및 가게번호 수집 )
    def crawl_coupang(self, driver, user_id, user_pw, row_idx):
        try:
            driver.get("https://store.coupangeats.com/merchant/management/stores/")
            time.sleep(5)
            
            # 로그인 처리
            id_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[1]/input')))
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(5)
            
            pw_input = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[2]/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(5)

            login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/button')))
            login_button.click()
            time.sleep(5)

            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 혹은 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.data_frame.at[row_idx, '연락처'] = "로그인 실패"
                driver.quit()
                return
            time.sleep(5)
            
            # "닫기" 버튼 클릭 처리, 화면에 표시된 버튼만 클릭
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                    time.sleep(5)
            # 연락처 값 크롤링
            try:
                contact_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div[1]/form/div/div/div[2]/div')
                self.scroll_to_element(driver, contact_element)
                contact_value = contact_element.text
                self.data_frame.at[row_idx, '연락처'] = contact_value
            except Exception as e:
                self.data_frame.at[row_idx, '연락처'] = f"에러 발생: {e}"
            time.sleep(5)
            # 사업자번호 값 크롤링
            try:
                business_number_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[1]/div[6]/div[5]/div')
                self.scroll_to_element(driver, business_number_element)
                business_number_value = business_number_element.text
                self.data_frame.at[row_idx, '사업자번호'] = business_number_value
            except Exception as e:
                self.data_frame.at[row_idx, '사업자번호'] = f"에러 발생: {e}"
            time.sleep(5)
            # 주소 값 크롤링
            try:
                address_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[1]/div[6]/div[12]/dl/dd[2]')
                self.scroll_to_element(driver, address_element)
                address_value = address_element.text
                self.data_frame.at[row_idx, '주소'] = address_value
            except Exception as e:
                self.data_frame.at[row_idx, '주소'] = f"에러 발생: {e}"
            time.sleep(5)
            # 사이트 이동 및 옵션 값 크롤링
            driver.get("https://self.baemin.com/")
            time.sleep(5)
            
            # "닫기" 버튼 클릭 처리, 화면에 표시된 버튼만 클릭
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                    time.sleep(5)
            try:
                select_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div[1]/div[4]/div[1]/div[1]/div/div/select')))
                options = select_element.find_elements(By.TAG_NAME, 'option')
                for idx, option in enumerate(options):
                    self.data_frame.at[row_idx, f'가게명_{idx+1}'] = option.text
            except Exception as e:
                self.data_frame.at[row_idx, '가게명'] = f"에러 발생: {e}"
            time.sleep(5)
        except Exception as e:
            print(f"쿠팡이츠 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"쿠팡이츠 에러 발생: {e}"

    def crawl_ddangyo(self, driver, user_id, user_pw, row_idx):
        try:
            driver.get("https://example.com/ddangyo")  # 땡겨요의 실제 URL 사용
            # 여기에 땡겨요의 크롤링 절차를 추가합니다.
            # 예시: 로그인, 페이지 이동, 데이터 수집 등
        except Exception as e:
            print(f"땡겨요 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"땡겨요 에러 발생: {e}"

# 전제 UI 활성화
if __name__ == '__main__':
    root = tk.Tk()
    root.title("크롤링 프로그램")
    root.geometry("1000x600")
    frame = Beamin_Crawling(root)
    frame.pack(fill="both", expand=True)
    root.mainloop()