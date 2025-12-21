import time
import datetime
from datetime import datetime
import pyautogui
import threading
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import Select
from undetected_chromedriver import Chrome, ChromeOptions
from tkinter import filedialog, messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    TimeoutException,
    NoSuchElementException,
)
from tkcalendar import DateEntry
import calendar


class CrawlingFinance(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.file_path = None
        self.data_frame = None
        self.progress_var = tk.DoubleVar()
        self.selected_service = tk.StringVar(value="배달의민족")
        self.start_date = None
        self.end_date = None

        # 다크 모드 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#2b2b2b')
        style.configure('TButton', background='#ffffff', foreground='#000000', font=('Helvetica', 12, 'bold'))
        style.configure('TLabel', background='#2b2b2b', foreground='#ffffff')
        style.configure('TProgressbar', foreground='#00ff00', background='#3c3c3c')
        style.configure('TProgressbar.Horizontal.TProgressbar', troughcolor='#3c3c3c', background='#00ff00')
        style.configure('Treeview', background='#2b2b2b', foreground='#ffffff', fieldbackground='#2b2b2b', bordercolor='#ffffff')
        style.map('Treeview', background=[('selected', '#3c3c3c')], foreground=[('selected', '#ffffff')])
        style.configure('Treeview.Heading', background='#3c3c3c', foreground='#ffffff')

        # 엑셀 파일 업로드 버튼 생성
        self.upload_button = ttk.Button(self, text="엑셀 파일 업로드", command=self.upload_file, style='TButton')
        self.upload_button.pack(pady=10)

        # 서비스 선택 프레임
        self.service_frame = ttk.Frame(self, style='TFrame')
        self.service_frame.pack(pady=10)
        self.baemin_radio = ttk.Radiobutton(self.service_frame, text="배달의민족", variable=self.selected_service, value="배달의민족", style='TRadiobutton')
        self.baemin_radio.pack(side=tk.LEFT)
        self.coupang_radio = ttk.Radiobutton(self.service_frame, text="쿠팡이츠", variable=self.selected_service, value="쿠팡이츠", style='TRadiobutton')
        self.coupang_radio.pack(side=tk.LEFT)
        self.yogiyo_radio = ttk.Radiobutton(self.service_frame, text="요기요", variable=self.selected_service, value="요기요", style='TRadiobutton')
        self.yogiyo_radio.pack(side=tk.LEFT)
        self.ddangyeo_radio = ttk.Radiobutton(self.service_frame, text="땡겨요", variable=self.selected_service, value="땡겨요", style='TRadiobutton')
        self.ddangyeo_radio.pack(side=tk.LEFT)

        # 날짜 선택 프레임
        self.date_frame = ttk.Frame(self, style='TFrame')
        self.date_frame.pack(pady=10)
        self.start_label = ttk.Label(self.date_frame, text="시작 날짜:", style='TLabel')
        self.start_label.pack(side=tk.LEFT, padx=5)
        self.start_date_entry = DateEntry(self.date_frame, date_pattern='yyyy-mm-dd', state='readonly')
        self.start_date_entry.pack(side=tk.LEFT, padx=5)
        self.end_label = ttk.Label(self.date_frame, text="종료 날짜:", style='TLabel')
        self.end_label.pack(side=tk.LEFT, padx=5)
        self.end_date_entry = DateEntry(self.date_frame, date_pattern='yyyy-mm-dd', state='readonly')
        self.end_date_entry.pack(side=tk.LEFT, padx=5)

        #시작 버튼 
        self.start_button = ttk.Button(self, text="시작", command=self.start_thread, style='TButton')
        self.start_button.pack(pady=10)

        # 진행률
        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, maximum=100, style='TProgressbar.Horizontal.TProgressbar')
        self.progress_bar.pack(pady=10, padx=10, fill=tk.X)
        self.progress_label = ttk.Label(self, text="진행 상황: 0%", style='TLabel')
        self.progress_label.pack(pady=10)

        self.treeview_frame = ttk.Frame(self, style='TFrame')
        self.treeview_frame.pack(fill="both", expand=True)
        self.treeview = ttk.Treeview(self.treeview_frame, style='Treeview')
        self.treeview.pack(fill="both", expand=True)

    # 화면상의 엑셀 상의 데이터를 삭제
    def clear_treeview(self):
        self.treeview.delete(*self.treeview.get_children())

    # 엑셀 업로드
    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        if self.file_path:
            self.data_frame = pd.read_excel(self.file_path)
            self.display_excel()

    # 엑셀 표시 (Treeview)
    def display_excel(self):
        if self.data_frame is not None:
            self.clear_treeview()
            self.treeview["column"] = list(self.data_frame.columns)
            self.treeview["show"] = "headings"
            for column in self.treeview["columns"]:
                self.treeview.heading(column, text=column)
            for row in self.data_frame.to_numpy().tolist():
                self.treeview.insert("", "end", values=row)

    # 진행중인 행 하이라이트
    def highlight_row(self, row_id):
        for item in self.treeview.get_children():
            self.treeview.item(item, tags="")
        self.treeview.item(row_id, tags=("highlight",))
        self.treeview.tag_configure("highlight", background="yellow", foreground="black")

    #스레드로 작업 분할
    def start_thread(self):
        threading.Thread(target=self.start_crawling).start()

    #크롤링 시작 시 예외처리 
    def start_crawling(self):
        if self.data_frame is not None:
            self.start_date = self.start_date_entry.get_date()
            self.end_date = self.end_date_entry.get_date()
            if self.start_date > self.end_date:
                messagebox.showwarning("경고", "시작 날짜는 종료 날짜 이전이어야 합니다.")
                return

            total_rows = len(self.data_frame.index)
            for index, row in self.data_frame.iterrows():
                item_id = self.treeview.get_children()[index]
                self.highlight_row(item_id)

                self.crawl_site(row["ID"], row["PW"], row["SHOP_ID"], index)

                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"진행 상황: {progress:.2f}%")
                self.update_idletasks()

            self.data_frame.to_excel(self.file_path, index=False)
            self.display_excel()
            messagebox.showinfo("완료", "완료됨~")
        else:
            messagebox.showwarning("경고", "엑셀파일부터 업로드 하셔야죠~")

    # 선택 옵션 별로 분리하기
    def crawl_site(self, user_id, user_pw, shop_id, row_idx):
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
            if selected_service == "배달의민족":
                self.crawl_baemin(driver, user_id, user_pw, shop_id, row_idx)
            elif selected_service == "쿠팡이츠":
                self.crawl_coupang(driver, user_id, user_pw, shop_id, row_idx)
            elif selected_service == "요기요":
                self.crawl_yogiyo(driver, user_id, user_pw, row_idx)
            elif selected_service == "땡겨요":
                self.crawl_ddangyeo(driver, user_id, user_pw, row_idx)
        except Exception as e:
            print(f"에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"에러 발생: {e}"
        finally:
            driver.quit()

    # 배민 DatePicker 헬퍼 (두 개 달력 지원, 왼/오른쪽 강제)
    def _parse_year_month(self, text: str):
        import re
        m = re.search(r'(\d{4})\s*년\s*(\d{1,2})\s*월', text)
        if not m:
            raise ValueError(f"연월 파싱 실패: {text}")
        return int(m.group(1)), int(m.group(2))

    #배민 사장님 사이트 상 날짜 지정 하는 캘린더 
    def _datepicker_root(self, driver):
        xp = '//*[@data-atelier-component="DatePicker" and @role="dialog" and @data-present="true"]'
        return WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xp)))

    # 배달의민족 달력 찾기
    def _get_calendars(self, driver):

        root = self._datepicker_root(driver)
        tables = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, './/table[@role="grid"]'))
        )
        cals = []
        for t in tables:
            cap = t.find_element(By.TAG_NAME, 'caption').text.strip()
            ym = self._parse_year_month(cap)
            cals.append((t, ym))
        if len(cals) < 2:
            raise RuntimeError("DatePicker에서 두 개의 달력을 찾지 못했습니다.")
        return cals  # [(table, (yyyy, mm)), (table, (yyyy, mm))]

    #배달의민족 달력에서 월 이동 
    def _shift_month(self, driver, direction: str):

        root = self._datepicker_root(driver)
        label = "다음 달" if direction == "next" else "이전 달"
        # 현재 보이는 월 기록
        before = [ym for _, ym in self._get_calendars(driver)]

        # 버튼은 모달 내에서 찾기
        btn = root.find_element(By.XPATH, f'.//button[@aria-label="{label}"]')
        try:
            WebDriverWait(driver, 5).until(lambda d: btn.is_displayed() and btn.is_enabled())
            btn.click()
        except Exception:
            # 폴백: JS 클릭
            driver.execute_script("arguments[0].click();", btn)

        # 월 변경 대기
        def months_changed(d):
            try:
                after = [ym for _, ym in self._get_calendars(d)]
                return after != before
            except Exception:
                return False

        WebDriverWait(driver, 10).until(months_changed)

    #배달의민족 목표 월로 이동 시키는 로직 
    def _ensure_month_on_side(self, driver, target_date, side="left"):
        side_idx = 0 if side == "left" else 1
        max_steps = 36

        def key(ym): return ym[0] * 12 + ym[1]
        tgt = (target_date.year, target_date.month)
        tgt_k = key(tgt)

        for _ in range(max_steps):
            cals = self._get_calendars(driver)
            left_ym = cals[0][1]
            right_ym = cals[1][1]
            left_k = key(left_ym)
            right_k = key(right_ym)

            # 원하는 패널이 이미 목표 월
            if cals[side_idx][1] == tgt:
                return cals[side_idx][0]

            # 목표가 현재 범위보다 앞/뒤인 경우 범위로 이동
            if tgt_k < left_k:
                self._shift_month(driver, "prev")
                continue
            if tgt_k > right_k:
                self._shift_month(driver, "next")
                continue

            # 목표 월이 화면에는 있으나 반대편 패널에만 있는 경우, 한 칸 이동해 원하는 패널로 보냄
            if left_ym == tgt and side_idx == 1:
                # [L, R]에서 L=tgt를 오른쪽으로 보내려면 prev → [L-1, L]
                self._shift_month(driver, "prev")
                continue
            if right_ym == tgt and side_idx == 0:
                # [L, R]에서 R=tgt를 왼쪽으로 보내려면 next → [R, R+1]
                self._shift_month(driver, "next")
                continue

        raise RuntimeError("목표 월을 원하는 패널에서 찾지 못했습니다.")

    #배달의민족 날짜 클릭 
    def _click_day_in_table(self, driver, table_el, date_obj):

        wait = WebDriverWait(driver, 10)
        d = int(str(date_obj.day))
        label_no0 = f"{d}일"
        label_02 = f"{d:02d}일"
        enabled_pred = '(not(@aria-disabled) or @aria-disabled="false")'

        candidates = [
            f'.//button[@aria-label="{label_no0}" and {enabled_pred}]',
            f'.//button[@aria-label="{label_02}" and {enabled_pred}]',
            f'.//button[{enabled_pred} and .//span[normalize-space(text())="{d}"]]',
            f'.//span[normalize-space(text())="{d}"]/ancestor::button[1][{enabled_pred}]',
        ]
        last_err = None
        for xp in candidates:
            try:
                btn = table_el.find_element(By.XPATH, xp)
                driver.execute_script("arguments[0].scrollIntoView({block:\"center\"});", btn)
                time.sleep(0.1)
                try:
                    wait.until(lambda d: btn.is_displayed() and btn.is_enabled())
                    btn.click()
                    return
                except Exception:
                    pass
                try:
                    ActionChains(driver).move_to_element(btn).pause(0.05).click().perform()
                    return
                except Exception as e2:
                    last_err = e2
                try:
                    driver.execute_script("arguments[0].click();", btn)
                    return
                except Exception as e3:
                    last_err = e3
            except (NoSuchElementException, StaleElementReferenceException) as e:
                last_err = e
                continue
        raise RuntimeError(f"일자 버튼 클릭 실패: {date_obj} / last_error={last_err}")

    def _clamp_to_month_last_day(self, date_obj):

        last_day = calendar.monthrange(date_obj.year, date_obj.month)[1]
        if date_obj.day > last_day:
            from datetime import date as _date
            return _date(date_obj.year, date_obj.month, last_day)
        return date_obj

    def set_baemin_date_range(self, driver):
        # 시작 월/일 (왼쪽)
        start_table = self._ensure_month_on_side(driver, self.start_date, side="left")
        self._click_day_in_table(driver, start_table, self.start_date)
        time.sleep(2)  # 요구사항

        # 종료 월/일 (필요 시 말일로 보정)
        safe_end = self._clamp_to_month_last_day(self.end_date)
        if (safe_end.year, safe_end.month) == (self.start_date.year, self.start_date.month):
            end_table = start_table  # 같은 월이면 왼쪽에서
            # 종료 일자 클릭 전에 화면을 조정하여 적용 버튼이 보이도록 스크롤
            apply_button_before_click = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//button[.//span[text()="적용"] and @data-atelier-component="Button"]'))
            )
            time.sleep(1)  # 스크롤 안정화 대기
        else:
            end_table = self._ensure_month_on_side(driver, safe_end, side="right")  # 다른 월이면 오른쪽에서

        self._click_day_in_table(driver, end_table, safe_end)
        time.sleep(3)

        # 1차 적용 버튼 
        # 취소 버튼 옆의 적용 버튼 찾기
        apply_button_xpath = '//span[text()="취소"]/ancestor::div//span[text()="적용"]/ancestor::button'
        apply_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, apply_button_xpath))
        )
        
        try:
            apply_button.click()
        except (ElementClickInterceptedException, StaleElementReferenceException):
            driver.execute_script("arguments[0].click();", apply_button)
            
        time.sleep(3)
        
        # 2차 적용 버튼을 찾을 때 동일하게 적용
        apply2_btn_xpath = '//button[.//span[text()="적용"] and @data-atelier-component="Button"]'
        apply2_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, apply2_btn_xpath))
        )
        try:
            apply2_btn.click()
            time.sleep(3)
        except (StaleElementReferenceException, ElementClickInterceptedException):
            apply2_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, apply2_btn_xpath))
            )
            driver.execute_script("arguments[0].click();", apply2_btn)
            
        # DatePicker 닫힘 대기(옵션)
        try:
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((
                    By.XPATH, '//*[@data-atelier-component="DatePicker" and @role="dialog" and @data-present="true"]'
                ))
            )
        except TimeoutException:
            pass

        # DatePicker 닫힘 대기(옵션)
        try:
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((
                    By.XPATH, '//*[@data-atelier-component="DatePicker" and @role="dialog" and @data-present="true"]'
                ))
            )
        except TimeoutException:
            pass

    # 배민 크롤링 작업
    def crawl_baemin(self, driver, user_id, user_pw, shop_id, row_idx):
        try:
            driver.get("https://self.baemin.com/orders/history")
            time.sleep(2)

            # ID 입력
            id_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[1]/div/div/form/div[1]/span/input'))
            )
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(1)

            # PW 입력
            pw_input = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div/form/div[2]/span/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(1)

            # 로그인 버튼 클릭
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div/form/button'))
            )
            login_button.click()
            time.sleep(2)

            # 로그인 실패 체크
            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 또는 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.data_frame.at[row_idx, '에러'] = "로그인 실패"
                return
            time.sleep(1)

            # 팝업 닫기
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    try:
                        WebDriverWait(driver, 3).until(lambda d: button.is_displayed() and button.is_enabled())
                        button.click()
                    except Exception:
                        try:
                            driver.execute_script("arguments[0].click();", button)
                        except Exception:
                            pass
                    time.sleep(0.3)

            # 날짜 직접 선택 버튼 클릭
            date_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div[1]/div[4]/div[1]/button[1]')))
            date_button.click()
            time.sleep(1)

            # 캘린더(Datepicker) 버튼 클릭
            try:
                date_calander_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[@data-atelier-component="DatePicker.Trigger"]'))
                )
                date_calander_button.click()
            except Exception as inner_error:
                print(f"캘린더 버튼 찾기 실패: {inner_error}")
                date_calander_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'DatePicker_b_r4ax_14nnus31'))
                )
                date_calander_button.click()
            time.sleep(1)

            # 시작/종료 날짜 적용 (왼/오 패널 규칙 + 월 이동 개선)
            self.set_baemin_date_range(driver)
            time.sleep(2)

            # 선택 화면 열기
            choose_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'FilterContainer-module__ccrG')))
            choose_button.click()
            time.sleep(2)
            
            # Shop 선택
            select_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select.Select_b_r4ax_11w1d6i7")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)  # 요소를 스크롤하여 중앙으로 이동
            time.sleep(1)  # 잠시 대기
            try:
                select_element.click()
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", select_element)

            # 옵션 값 선택
            option_xpath = f'//select[contains(@class, "Select_b_r4ax_11w1d6i7")]//option[@value="{shop_id}"]'
            option_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option_element)
            time.sleep(1)  # 잠시 대기
            try:
                option_element.click()
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", option_element)
            time.sleep(2)

            # 적용 버튼 클릭
            apply_button_xpath = '//span[text()="적용"]/ancestor::button[@aria-disabled="false"]'
            apply_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, apply_button_xpath)))
            try:
                apply_button.click()
            except (ElementClickInterceptedException, StaleElementReferenceException):
                driver.execute_script("arguments[0].click();", apply_button)
            time.sleep(2)

            # 매출금액 수집
            try:
                finance_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "TotalSummary-module__OYdy")]/div[2]/span[contains(@class, "TotalSummary-module__SysK")]/b'))
                )
                finance_value = finance_element.text.replace(",", "").replace("원", "").strip()
                self.data_frame.at[row_idx, '매출금액'] = finance_value
            except Exception as e:
                self.data_frame.at[row_idx, '매출금액'] = f"에러 발생: {e}"
            time.sleep(1)

            # 매출건수 수집
            try:
                finance_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "TotalSummary-module__OYdy")]/div[1]/span[contains(@class, "TotalSummary-module__SysK")]/b'))
                )
                finance_value = finance_element.text.replace(",", "").replace("건", "").strip()
                self.data_frame.at[row_idx, '매출건수'] = finance_value
            except Exception as e:
                self.data_frame.at[row_idx, '매출건수'] = f"에러 발생: {e}"
            time.sleep(1)

        except Exception as e:
            print(f"배달의민족 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"배달의민족 에러 발생: {e}"


    #쿠팡이츠 크롤링 작업
    def crawl_coupang(self, driver, user_id, user_pw, shop_id, row_idx):
        try:
            driver.get("https://store.coupangeats.com/merchant/login")
            time.sleep(2)
            
            # ID 입력
            id_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[1]/input')))
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(1)
            
            # PW 입력
            pw_input = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[2]/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(1)
            
            # 로그인 버튼 클릭 
            login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/button')))
            login_button.click()
            time.sleep(3)
            
            # 로그인 실패 시 C 열에 로그인 실패 기록 후 종료 
            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 혹은 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.data_frame.at[row_idx, '에러'] = "로그인 실패"
                driver.quit()
                return
            time.sleep(2)

            # 사이트 이동 ( 가맹점 화면으로 이동 )
            driver.get(f"https://store.coupangeats.com/merchant/management/orders/{shop_id}")
            time.sleep(2)
            
            popup_xpaths = [
                '/html/body/div[4]/div/div/div/button',
                '/html/body/div[3]/div/div/div/button',
                '/html/body/div[3]/div/div/div/div[2]/button[2]'
            ]

            def close_popups(driver, max_round=5):
                for _ in range(max_round):
                    closed_any = False

                    for xpath in popup_xpaths:
                        try:
                            popup = WebDriverWait(driver, 1).until(
                                EC.element_to_be_clickable((By.XPATH, xpath))
                            )
                            popup.click()
                            time.sleep(0.4)
                            closed_any = True

                        except (TimeoutException, StaleElementReferenceException, ElementClickInterceptedException):
                            try:
                                popup = WebDriverWait(driver, 1).until(
                                    EC.presence_of_element_located((By.XPATH, xpath))
                                )
                                driver.execute_script("arguments[0].click();", popup)
                                time.sleep(0.4)
                                closed_any = True
                            except:
                                pass

                    if not closed_any:
                        break

            #팝업 끄기 실행 
            close_popups(driver)
            time.sleep(2)

            # 위치 기준 클릭을 통해 팝업 창 닫기 
            action = ActionChains(driver)
            action.move_by_offset(1, 1).click().perform()
            time.sleep(2)
            
            # 날짜 지정 버튼 클릭 
            Date_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div/div/div/div/div[2]/div[2]/div/div[1]')))
            Date_button.click()
            time.sleep(2)
            
            # 달력 버튼 클릭 (달력 표시)
            Date_Calander_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div/div/div/div/div[2]/div[2]/div/div[2]/div[1]/div[1]/div')))
            Date_Calander_button.click()
            time.sleep(2)
            
            # 시작일자 및 종료일자 설정
            start_date = self.start_date
            end_date = self.end_date

            # 시작 날짜 달력 요소 선택 및 클릭
            self.select_date(driver, start_date)

            # 달력 버튼 클릭 (달력 표시)
            Date_Calander_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div/div/div/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div')))
            Date_Calander_button.click()
            time.sleep(2)

            # 종료 날짜 달력 요소 선택 및 클릭
            self.select_date(driver, end_date)

            # 조회 버튼 클릭 
            apply_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div/div/div/div/div[2]/div[2]/button')))
            apply_button.click()
            time.sleep(2)

            # 매출금액 수집
            try:
                finance_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div/div/div/div/div[3]/div/div[1]/div[2]/span[1]'))
                )
                finance_value = finance_element.text.replace(",", "").replace("원", "").strip()
                self.data_frame.at[row_idx, '매출금액'] = finance_value
            except Exception as e:
                self.data_frame.at[row_idx, '매출금액'] = f"에러 발생: {e}"
            time.sleep(2)

            # 매출건수 수집
            try:
                finance_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div/div/div/div/div[3]/div/div[2]/div[2]/span[1]'))
                )
                finance_value = finance_element.text.replace(",", "").replace("건", "").strip()
                self.data_frame.at[row_idx, '매출건수'] = finance_value
            except Exception as e:
                self.data_frame.at[row_idx, '매출건수'] = f"에러 발생: {e}"
            time.sleep(2)


        except Exception as e:
            print(f"쿠팡이츠 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"쿠팡이츠 에러 발생: {e}"

    def select_date(self, driver, target_date):
        day = target_date.day
        month = target_date.strftime('%B')
        year = target_date.year

        # 달력의 월과 년도 선택 - 타겟 달로 이동
        while True:
            month_year_display = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "DayPicker-Caption")]'))
            ).text
            current_month, current_year = month_year_display.split(' ')

            # 목표 연월과 현재 달력의 연월 비교
            target_date_obj = datetime.strptime(f'{month} {year}', '%B %Y')
            current_date_obj = datetime.strptime(f'{current_month} {current_year}', '%B %Y')

            if current_date_obj == target_date_obj:
                break
            elif current_date_obj < target_date_obj:
                # 목표 달이 현재 달보다 미래인 경우
                next_button = driver.find_element(By.XPATH, '//span[contains(@class, "DayPicker-NavButton--next")]')
                driver.execute_script("arguments[0].click();", next_button)
            else:
                # 목표 달이 현재 달보다 과거인 경우
                prev_button = driver.find_element(By.XPATH, '//span[contains(@class, "DayPicker-NavButton--prev")]')
                driver.execute_script("arguments[0].click();", prev_button)

            time.sleep(1)

        # 타겟 날짜 선택
        day_element_xpath = f'//div[contains(@class, "DayPicker-Day") and not(contains(@class, "DayPicker-Day--outside")) and text()="{day}"]'
        day_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, day_element_xpath))
        )
        driver.execute_script("arguments[0].click();", day_element)
        time.sleep(1)


    # 요기요 크롤링 작업
    def crawl_yogiyo(self, driver, user_id, user_pw, row_idx):
        try:
            driver.get("https://example.com/yogiyo")
            time.sleep(3)
        except Exception as e:
            print(f"요기요 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"요기요 에러 발생: {e}"

    # 땡겨요 크롤링 작업
    def crawl_ddangyeo(self, driver, user_id, user_pw, row_idx):
        try:
            driver.get("https://boss.ddangyo.com/")
            time.sleep(2)
            
            # ID 입력
            id_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/ul/li[1]/div[2]/div/input')))
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(1)
            
            # PW 입력
            pw_input = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/ul/li[2]/div[2]/div/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(1)
            
            # 로그인 버튼 클릭 
            login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/input')))
            login_button.click()
            time.sleep(2)

            # 로그인 실패 체크
            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 또는 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.data_frame.at[row_idx, '에러'] = "로그인 실패"
                return
            time.sleep(1)

            #비밀번호 변경 확인 
            try:
                pwchange_button = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH,'/html/body/div[4]/div[2]/div[1]/div/div[2]/input[1]')))
                pwchange_button.click()
                time.sleep(1)

            except TimeoutException:
                pass  # 비밀번호 변경 팝업 없음 → 정상 진행
            
            time.sleep(1)

            # 팝업 닫기
            close_buttons = driver.find_elements(By.XPATH, '//input[@value="닫기"]')
            for button in close_buttons:
                if button.is_displayed():
                    try:
                        WebDriverWait(driver, 10).until(lambda d: button.is_displayed() and button.is_enabled())
                        button.click()
                    except Exception:
                        try:
                            driver.execute_script("arguments[0].click();", button)
                        except Exception:
                            pass
                    time.sleep(1)

            # 주문내역 버튼 클릭 
            list_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'//a[contains(@class,"w2anchor") and normalize-space(text())="주문내역"]')))
            list_button.click()
            time.sleep(1)

            # 시작 날짜 입력 (YYYYMMDD)
            start_date_str = self.start_date.strftime("%Y%m%d")
            start_date_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[3]/div[1]/section/div/div[1]/fieldset/div[1]/div[3]/div[1]/div/div[1]/input')))
            start_date_input.clear()
            start_date_input.send_keys(start_date_str)
            time.sleep(1)

            # 종료 날짜 입력 (YYYYMMDD)
            end_date_str = self.end_date.strftime("%Y%m%d")
            end_date_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[3]/div[1]/section/div/div[1]/fieldset/div[1]/div[3]/div[3]/div/div[1]/input')))
            end_date_input.clear()
            end_date_input.send_keys(end_date_str)
            time.sleep(1)
            
            #조회 버튼 클릭 
            apply_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/section/div/div[1]/fieldset/div[2]/input')))
            apply_button.click()
            time.sleep(2)

            # 조회 결과 없음 체크
            no_data_elements = driver.find_elements(By.XPATH,'//strong[contains(text(),"조회된 내역이 없습니다")]')
            if no_data_elements and no_data_elements[0].is_displayed():
                self.data_frame.at[row_idx, '매출금액'] = 0
                self.data_frame.at[row_idx, '매출건수'] = 0
                return  

            # 매출금액 수집
            try:
                finance_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[3]/div[1]/section/div/div[2]/div[2]/dl/dd/span[1]'))
                )
                finance_value = finance_element.text.replace(",", "").replace("원", "").strip()
                self.data_frame.at[row_idx, '매출금액'] = finance_value
            except Exception as e:
                self.data_frame.at[row_idx, '매출금액'] = f"에러 발생: {e}"
            time.sleep(1)

            # 매출건수 수집
            try:
                finance_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[3]/div[1]/section/div/div[2]/div[1]/dl/dd/span[1]'))
                )
                finance_value = finance_element.text.replace(",", "").replace("건", "").strip()
                self.data_frame.at[row_idx, '매출건수'] = finance_value
            except Exception as e:
                self.data_frame.at[row_idx, '매출건수'] = f"에러 발생: {e}"
            time.sleep(1)


        except Exception as e:
            print(f"땡겨요 에러 발생: {e}")
            self.data_frame.at[row_idx, '에러'] = f"땡겨요 에러 발생: {e}"


# Initialize UI
if __name__ == '__main__':
    root = tk.Tk()
    root.title("크롤링 프로그램")
    root.geometry("1000x600")
    frame = CrawlingFinance(root)
    frame.pack(fill="both", expand=True)
    root.mainloop()