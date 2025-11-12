import time
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

        # Readonly 입력칸 / 검색 버튼들 프레임 생성
        self.input_frame = ttk.Frame(self, style='TFrame')
        self.input_frame.pack(pady=0, padx=0)
        
        # Readonly 입력칸 / 검색 버튼들 프레임 생성
        self.input_frame = ttk.Frame(self, style='TFrame')
        self.input_frame.pack(pady=0, padx=0)

        # 배달의민족 Class 값을 받아올 label 생성
        self.label1 = ttk.Label(self.input_frame, text="배달의민족 Class 확인 필요", style='TLabel', background='#2b2b2b', foreground='#ffffff', anchor='w')
        self.label1.pack(side=tk.LEFT, pady=0, padx=10, fill=tk.X, expand=True)

        #배달의민족 Class 검색 버튼 
        self.search_button1 = ttk.Button(self.input_frame, text="검색", command=self.search_function1, style='TButton')
        self.search_button1.pack(side=tk.LEFT, pady=0, padx=10)

        # 쿠팡이츠 Class 값을 받아올 label 생성
        self.label2 = ttk.Label(self.input_frame, text="쿠팡이츠 Class 확인 필요", style='TLabel', background='#2b2b2b', foreground='#ffffff', anchor='w')
        self.label2.pack(side=tk.LEFT, pady=0, padx=10, fill=tk.X, expand=True)

        #쿠팡이츠 Class 검색 버튼
        self.search_button2 = ttk.Button(self.input_frame, text="검색", command=self.search_function2, style='TButton')
        self.search_button2.pack(side=tk.LEFT, pady=0, padx=0)

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
            self.display_excel()

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


    def search_function1(self):
        print("Search function 1 executed")

    def search_function2(self):
        print("Search function 2 executed")

    def start_thread(self):
        threading.Thread(target=self.start_crawling).start()

    def start_crawling(self):
        if self.data_frame is not None:
            total_rows = len(self.data_frame.index)
            for index, row in self.data_frame.iterrows():
                item_id = self.treeview.get_children()[index]  # 해당 행의 ID
                self.highlight_row(item_id)  # 현재 행을 하이라이트
                
                self.crawl_site(row["ID"], row["PW"], index)
                
                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"진행 상황: {progress:.2f}%")
                self.update_idletasks()
            
            #데이터를 엑셀에 저장 
            self.data_frame.to_excel(self.file_path, index=False)

            #데이터 저장 후 화면에 다시 표출
            self.display_excel()

            messagebox.showinfo("완료", "완료됨~")
        else:
            messagebox.showwarning("경고", "엑셀파일부터 업로드 하셔야죠~")

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
        
        driver = Chrome(options=options, use_subprocess=True)
        try:
            # 크롤링 작업 선택 ( 배달의민족 / 쿠팡이츠 / 땡겨요 )
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

    # 배민 크롤링 작업
    def crawl_baemin(self, driver, user_id, user_pw, row_idx):
        try:
            #배민 사이트 접속 
            driver.get("https://self.baemin.com/mypage/owner")
            time.sleep(5)
            
            #ID 입력
            id_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[1]/div/div/form/div[1]/span/input')))
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(5)

            #PW 입력
            pw_input = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/form/div[2]/span/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(5)

            #로그인 버튼 클릭
            login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div/form/button')))
            login_button.click()
            time.sleep(5)

            #로그인 실패 시 C열에 로그인 실패 기록 후 종료 
            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 또는 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.data_frame.at[row_idx, '연락처'] = "로그인 실패"
                driver.quit()
                return
            time.sleep(5)
            
            #팝업 창 닫기 
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                    time.sleep(5)
            
            # 연락처 정보 수집 
            try:
                contact_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div[1]/form/div/div/div[2]/div')
                self.scroll_to_element(driver, contact_element)
                contact_value = contact_element.text
                self.data_frame.at[row_idx, '연락처'] = contact_value
            except Exception as e:
                self.data_frame.at[row_idx, '연락처'] = f"에러 발생: {e}"
            time.sleep(5)
            
            # 사업자번호 수집
            try:
                business_number_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[1]/div[6]/div[5]/div')
                self.scroll_to_element(driver, business_number_element)
                business_number_value = business_number_element.text
                self.data_frame.at[row_idx, '사업자번호'] = business_number_value
            except Exception as e:
                self.data_frame.at[row_idx, '사업자번호'] = f"에러 발생: {e}"
            time.sleep(5)
            
            # 주소 수집 
            try:
                address_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div[1]/div[6]/div[12]/dl/dd[2]')
                self.scroll_to_element(driver, address_element)
                address_value = address_element.text
                self.data_frame.at[row_idx, '주소'] = address_value
            except Exception as e:
                self.data_frame.at[row_idx, '주소'] = f"에러 발생: {e}"
            time.sleep(5)
            
            #사이트 이동 ( 홈 화면으로 이동 )
            driver.get("https://self.baemin.com/")
            time.sleep(5)
            
            #팝업 창 닫기 
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                    time.sleep(5)
            
            # 가맹점 정보 수집 
            try:
                select_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div[1]/div[4]/div[1]/div[1]/div/div/select')))
                options = select_element.find_elements(By.TAG_NAME, 'option')
                for idx, option in enumerate(options):
                    self.data_frame.at[row_idx, f'가게명_{idx+1}'] = option.text
            except Exception as e:
                self.data_frame.at[row_idx, '가게명'] = f"에러 발생: {e}"
            time.sleep(5)
            
            driver.delete_all_cookies()
        except Exception as e:
            print(f"배달의민족 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"배달의민족 에러 발생: {e}"

    # 쿠팡이츠 크롤링 작업 
    def crawl_coupang(self, driver, user_id, user_pw, row_idx):
        try:
            driver.get("https://store.coupangeats.com/merchant/management/stores/")
            time.sleep(5)
            
            #ID 입력
            id_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[1]/input')))
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(5)
            
            #PW 입력
            pw_input = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[2]/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(5)

            #로그인 버튼 클릭 
            login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/button')))
            login_button.click()
            time.sleep(5)

            #로그인 실패 시 C열에 로그인 실패 기록 후 종료 
            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 혹은 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.data_frame.at[row_idx, '연락처'] = "로그인 실패"
                driver.quit()
                return
            time.sleep(5)
            
            #사이트 이동 ( 가맹점 화면으로 이동 )
            driver.get("https://store.coupangeats.com/merchant/management/stores/")
            time.sleep(5)

            # 클릭할 위치에 마우스 이동 후 클릭
            action = ActionChains(driver)
            action.move_by_offset(10, 10).click().perform()
            time.sleep(5)

            #다음에 하기 버튼 클릭 
            close_buttons = driver.find_elements(By.XPATH, '//button[text()="다음에 하기"]')
            for button in close_buttons:
                if button.is_displayed():
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button)).click()
                    time.sleep(5)

            # 가게 정보 수집 
            try:
                # 가게 정보 요소 찾기
                store_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'store-link')))
                
                for idx, store_element in enumerate(store_elements):
                    # Debugging: 요소의 HTML 내용을 출력
                    print(f"Store element HTML content: {store_element.get_attribute('innerHTML')}")

                    # 가게명 저장
                    store_name = store_element.text
                    self.data_frame.at[row_idx, f'가게명_{idx+1}'] = store_name
                    
                    # 가게번호 저장 (바로 다음 첫 번째 <span> 요소)
                    try:
                        # store_element의 부모 요소에서 첫 번째 <span>을 찾기
                        parent_element = store_element.find_element(By.XPATH, "..")  # ".."으로 상위 요소 접근
                        first_span_element = parent_element.find_element(By.XPATH, './span[1]')
                        
                        if first_span_element: # 첫 번째 <span> 요소가 존재한다면
                            store_number = first_span_element.text
                            self.data_frame.at[row_idx, f'가게번호_{idx+1}'] = store_number
                        else:
                            self.data_frame.at[row_idx, f'가게번호_{idx+1}'] = 'No span found'
                    except Exception as e:
                        self.data_frame.at[row_idx, f'가게번호_{idx+1}'] = f"에러 발생: {e}"
            except Exception as e:
                self.data_frame.at[row_idx, '가게명'] = f"에러 발생: {e}"
            time.sleep(5)

        except Exception as e:
            print(f"쿠팡이츠 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"쿠팡이츠 에러 발생: {e}"


# Initialize UI
if __name__ == '__main__':
    root = tk.Tk()
    root.title("크롤링 프로그램")
    root.geometry("1000x600")
    frame = Takedown(root)
    frame.pack(fill="both", expand=True)
    root.mainloop()