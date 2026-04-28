import re
import os
import winreg
import time
import json
import threading
import logging
import traceback
import webbrowser
from tkinter import messagebox
from io import BytesIO

import pandas as pd
import requests
import tkinter as tk
from tkinter import ttk, messagebox

import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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


class MultiCrawlerApp(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.data_frame = None
        self.is_running = False
        self.progress_var = tk.DoubleVar()
        self.current_status_var = tk.StringVar(value="대기 중")

        # 설정값
        self.google_xlsx_url = "https://docs.google.com/spreadsheets/d/1iCMikJwc3FqxNtly8zkh6GFdX4luvs155TH2mbsVxd0/export?format=xlsx"
        self.source_sheet_name = "List"
        self.result_sheet_map = {
        "배달의민족": "배달의민족_결과",
        "요기요": "요기요_결과",
        "쿠팡이츠": "쿠팡이츠_결과",
        "땡겨요": "땡겨요_결과"
        }
        self.apps_script_url = "https://script.google.com/macros/s/AKfycbzO3Nnt3Spnhj_JCcsfkPGrWpjazDR5vOrJC31CZpcSiMfNi9FwqMpphDpUT_lfe9xnqg/exec"

        # 서비스별 크롤링 설정
        self.crawl_settings = {
            "배달의민족": {
                "가맹점명": True,
                "샵넘버": True,
                "광고서비스사용유무": True,
                "주문접수채널": True,
                "연락처": True,
                "사업자번호": True,
                "주소": True
            },
            "요기요": {
                "가맹점명": True,
                "샵넘버": True,
                "연락처": True,
                "사업자번호": True,
                "제휴포스사용여부": True
            },
            "쿠팡이츠": {
                "가게명": True,
                "샵넘버": True,
                "상태": True
            },
            "땡겨요": {
                "가게명": True,
                "샵넘버": True
            }
        }

        self.setting_vars = {}

        self.init_logging()
        self.init_style()
        self.create_widgets()

    # --------------------------
    # 초기화
    # --------------------------
    def init_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s"
        )

    def init_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TFrame", background="#2b2b2b")
        style.configure("TLabel", background="#2b2b2b", foreground="#ffffff")
        style.configure("TButton", background="#ffffff", foreground="#000000", font=("Helvetica", 10, "bold"))
        style.configure("Treeview", background="#2b2b2b", foreground="#ffffff", fieldbackground="#2b2b2b")
        style.configure("Treeview.Heading", background="#3c3c3c", foreground="#ffffff")
        style.configure("TProgressbar.Horizontal.TProgressbar",
                        troughcolor="#3c3c3c", background="#00ff00")

    def create_widgets(self):
        self.configure(bg="#2b2b2b")

        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=10, pady=10)

        self.load_button = ttk.Button(top_frame, text="구글 시트 불러오기", command=self.load_sheet_from_google)
        self.load_button.pack(side="left", padx=5)

        self.setting_button = ttk.Button(top_frame, text="설정", command=self.open_settings_window)
        self.setting_button.pack(side="left", padx=5)

        self.start_button = ttk.Button(top_frame, text="시작", command=self.start_thread)
        self.start_button.pack(side="left", padx=5)

        self.retry_button = ttk.Button(top_frame, text="결과 재전송", command=self.retry_upload_results)
        self.retry_button.pack(side="left", padx=5)

        self.retry_button = ttk.Button(top_frame, text="시트 확인", command=self.open_googlesheet)
        self.retry_button.pack(side="left", padx=5)

        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, maximum=100,
                                            style="TProgressbar.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", padx=10, pady=5)

        self.progress_label = ttk.Label(self, text="진행 상황: 0%")
        self.progress_label.pack(anchor="w", padx=10)

        self.current_status_label = ttk.Label(self, textvariable=self.current_status_var)
        self.current_status_label.pack(anchor="w", padx=10, pady=(0, 10))

        self.treeview_frame = ttk.Frame(self)
        self.treeview_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.treeview = ttk.Treeview(self.treeview_frame)
        self.treeview.pack(fill="both", expand=True)

        log_frame = ttk.Frame(self)
        log_frame.pack(fill="both", padx=10, pady=10)

        ttk.Label(log_frame, text="실행 로그").pack(anchor="w")

        self.log_text = tk.Text(log_frame, height=10, bg="#1e1e1e", fg="#ffffff", insertbackground="#ffffff")
        self.log_text.pack(fill="both", expand=True)

    # --------------------------
    # UI 유틸
    # --------------------------
    def safe_ui(self, func, *args, **kwargs):
        self.after(0, lambda: func(*args, **kwargs))

    def write_log(self, message):
        logging.info(message)

        def _append():
            self.log_text.insert(tk.END, f"{message}\n")
            self.log_text.see(tk.END)

        self.safe_ui(_append)

    def clear_treeview(self):
        self.treeview.delete(*self.treeview.get_children())

    def display_data(self):
        if self.data_frame is None:
            return

        self.clear_treeview()
        self.treeview["columns"] = list(self.data_frame.columns)
        self.treeview["show"] = "headings"

        for col in self.treeview["columns"]:
            self.treeview.heading(col, text=col)
            self.treeview.column(col, width=120, anchor="center")

        for row in self.data_frame.fillna("").to_numpy().tolist():
            self.treeview.insert("", "end", values=row)

    def refresh_treeview_row(self, row_idx):
        children = self.treeview.get_children()
        if row_idx >= len(children):
            return
        values = self.data_frame.iloc[row_idx].fillna("").tolist()
        self.treeview.item(children[row_idx], values=values)

    def highlight_row(self, row_idx):
        children = self.treeview.get_children()
        for item in children:
            self.treeview.item(item, tags=())
        if 0 <= row_idx < len(children):
            self.treeview.item(children[row_idx], tags=("highlight",))
            self.treeview.tag_configure("highlight", background="yellow", foreground="black")

    # --------------------------
    # 설정창
    # --------------------------
    def open_settings_window(self):
        settings_window = tk.Toplevel(self)
        settings_window.title("크롤링 설정")
        settings_window.geometry("500x500")
        settings_window.configure(bg="#2b2b2b")
        settings_window.transient(self.winfo_toplevel())
        settings_window.grab_set()

        notebook = ttk.Notebook(settings_window)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.setting_vars = {}

        for service, items in self.crawl_settings.items():
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=service)

            self.setting_vars[service] = {}

            row_num = 0
            for item_name, enabled in items.items():
                var = tk.BooleanVar(value=enabled)
                self.setting_vars[service][item_name] = var

                chk = tk.Checkbutton(
                    frame,
                    text=item_name,
                    variable=var,
                    bg="#2b2b2b",
                    fg="#ffffff",
                    selectcolor="#3c3c3c",
                    activebackground="#2b2b2b",
                    activeforeground="#ffffff"
                )
                chk.grid(row=row_num, column=0, sticky="w", padx=20, pady=8)
                row_num += 1

        btn_frame = ttk.Frame(settings_window)
        btn_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(btn_frame, text="저장", command=lambda: self.save_settings(settings_window)).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="취소", command=settings_window.destroy).pack(side="right", padx=5)

    def save_settings(self, window):
        for service, items in self.setting_vars.items():
            for item_name, var in items.items():
                self.crawl_settings[service][item_name] = var.get()

        self.write_log(f"설정 저장 완료: {json.dumps(self.crawl_settings, ensure_ascii=False)}")
        messagebox.showinfo("완료", "설정이 저장되었습니다.")
        window.destroy()

    # --------------------------
    # 데이터 로드
    # --------------------------
    def load_sheet_from_google(self):
        try:
            self.write_log("구글 스프레드시트 다운로드 시작")
            self.current_status_var.set("구글 시트 다운로드 중...")

            response = requests.get(self.google_xlsx_url, timeout=30)
            response.raise_for_status()

            excel_data = BytesIO(response.content)
            self.data_frame = pd.read_excel(excel_data, sheet_name=self.source_sheet_name)
            self.data_frame.columns = self.data_frame.columns.astype(str).str.strip()

            self.ensure_columns()
            self.display_data()

            self.current_status_var.set("불러오기 완료")
            self.write_log(f"{self.source_sheet_name} 시트 로드 완료 - {len(self.data_frame)}건")

        except Exception as e:
            self.write_log(f"구글 시트 로드 실패: {e}")

    def ensure_columns(self):
        if self.data_frame is None:
            return

        required_columns = [
        "구분", "ID", "PW",
        "상태", "작업단계", "에러", "처리시간", "처리일시",
        "주문접수채널", "연락처", "사업자번호", "주소",
        "지번주소", "상세주소", "우편번호"
    ]

        for col in required_columns:
            if col not in self.data_frame.columns:
                self.data_frame[col] = ""

    # --------------------------
    # 업로드
    # --------------------------
    def upload_results_to_apps_script(self):
        try:
            if self.data_frame is None or self.data_frame.empty:
                messagebox.showwarning("경고", "업로드할 데이터가 없습니다.")
                return

            upload_df = self.data_frame.copy().fillna("")

            if "구분" not in upload_df.columns:
                raise Exception("업로드 데이터에 '구분' 컬럼이 없습니다.")

            grouped = upload_df.groupby("구분", dropna=False)

            success_count = 0
            fail_count = 0

            for service_type, group_df in grouped:
                service_type = str(service_type).strip()

                if not service_type:
                    self.write_log("구분값이 비어 있는 데이터는 업로드하지 않습니다.")
                    continue

                sheet_name = self.result_sheet_map.get(service_type)
                if not sheet_name:
                    self.write_log(f"구분값 '{service_type}'에 대한 결과 시트 매핑이 없습니다.")
                    fail_count += 1
                    continue

                payload = {
                    "sheetName": sheet_name,
                    "mode": "append",
                    "columns": list(group_df.columns),
                    "rows": group_df.astype(str).values.tolist()
                }

                self.write_log(f"[업로드 시작] {service_type} → {sheet_name} / {len(group_df)}건")

                try:
                    response = requests.post(self.apps_script_url, json=payload, timeout=60)
                    response.raise_for_status()

                    result = response.json()
                    if result.get("success"):
                        self.write_log(f"[업로드 성공] {sheet_name} / {len(group_df)}건")
                        success_count += 1
                    else:
                        msg = result.get("message", "알 수 없는 오류")
                        self.write_log(f"[업로드 실패] {sheet_name} / {msg}")
                        fail_count += 1

                except Exception as sub_e:
                    self.write_log(f"[업로드 오류] {sheet_name} / {sub_e}")
                    fail_count += 1

            if fail_count == 0:
                messagebox.showinfo("업로드 완료", f"모든 결과 업로드 완료 ({success_count}개 탭)")
            else:
                messagebox.showwarning("업로드 일부 실패", f"성공: {success_count} / 실패: {fail_count}")

        except Exception as e:
            self.write_log(f"결과 업로드 오류: {e}")
            messagebox.showerror("업로드 오류", f"Apps Script 업로드 실패:\n{e}")

    def retry_upload_results(self):
        if self.data_frame is None:
            messagebox.showwarning("경고", "재전송할 데이터가 없습니다.")
            return

        if not messagebox.askyesno("확인", "현재 결과를 다시 업로드하시겠습니까?"):
            return

        self.upload_results_to_apps_script()

    # --------------------------
    # 실행
    # --------------------------
    def start_thread(self):
        if self.is_running:
            messagebox.showwarning("안내", "이미 실행 중입니다.")
            return

        if self.data_frame is None or self.data_frame.empty:
            messagebox.showwarning("경고", "먼저 데이터를 불러와주세요.")
            return

        self.is_running = True
        self.start_button.config(state="disabled")
        threading.Thread(target=self.start_crawling, daemon=True).start()

    def start_crawling(self):
        try:
            total_rows = len(self.data_frame.index)
            self.write_log("전체 크롤링 시작")

            for index, row in self.data_frame.iterrows():
                service_type = str(row.get("구분", "")).strip()
                self.safe_ui(self.highlight_row, index)
                self.safe_ui(self.current_status_var.set, f"{index + 1}/{total_rows} 처리 중 - {service_type}")

                start_time = time.time()
                self.write_log(f"[{index + 1}/{total_rows}] 시작 - {service_type}")

                try:
                    self.process_row(row, index)
                except Exception as e:
                    self.set_row_result(index, status="실패", step="행 처리", error=str(e))
                    self.write_log(f"[행 {index + 1}] 예외 발생: {e}")

                elapsed = round(time.time() - start_time, 2)
                self.data_frame.at[index, "처리시간"] = elapsed
                self.data_frame.at[index, "처리일시"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

                progress = ((index + 1) / total_rows) * 100
                self.safe_ui(self.progress_var.set, progress)
                self.safe_ui(self.progress_label.config, text=f"진행 상황: {progress:.2f}%")
                self.safe_ui(self.refresh_treeview_row, index)

            self.safe_ui(self.current_status_var.set, "결과 업로드 중...")
            self.upload_results_to_apps_script()

            self.safe_ui(self.current_status_var.set, "작업 완료")
            self.write_log("전체 작업 완료")
            self.safe_ui(messagebox.showinfo, "완료", "전체 작업이 완료되었습니다.")

        except Exception as e:
            self.write_log(f"전체 실행 오류: {e}\n{traceback.format_exc()}")
            self.safe_ui(messagebox.showerror, "오류", f"실행 중 오류 발생:\n{e}")
        finally:
            self.is_running = False
            self.safe_ui(self.start_button.config, state="normal")

    # --------------------------
    # 행 처리 분기
    # --------------------------
    def process_row(self, row, row_idx):
        service_type = str(row.get("구분", "")).strip()

        if not service_type:
            self.set_row_result(row_idx, status="실패", step="구분 확인", error="구분값이 비어 있습니다.")
            return

        if service_type == "쿠팡이츠":
            driver = self.create_coupang_driver()
        else:
            driver = self.create_driver()

        try:
            if service_type == "배달의민족":
                self.process_baemin(driver, row, row_idx)
            elif service_type == "요기요":
                self.process_yogiyo(driver, row, row_idx)
            elif service_type == "쿠팡이츠":
                self.process_coupang(driver, row, row_idx)
            elif service_type == "땡겨요":
                self.process_ddangyo(driver, row, row_idx)
            else:
                self.set_row_result(row_idx, status="실패", step="구분 분기", error=f"지원하지 않는 구분값: {service_type}")
        finally:
            try:
                driver.quit()
            except:
                pass

    # --------------------------
    # 공통 Chromedriver  
    # --------------------------
    def create_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-plugins-discovery")
        options.add_argument("--incognito")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--log-level=3")

        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options=options)
        driver.set_window_size(1400, 900)
        return driver

    # --------------------------
    # 쿠팡이츠용 Undetected Chromedriver  
    # --------------------------  
    def create_coupang_driver(self):
        chrome_major_version = self.get_chrome_major_version()

        options = uc.ChromeOptions()
        options.add_argument("--start-maximized")

        driver = uc.Chrome(
            options=options,
            use_subprocess=True,
            version_main=chrome_major_version
        )
        driver.set_window_size(1400, 900)

        try:
            driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {
                    "source": """
                        Object.defineProperty(navigator, 'webdriver', {
                            get: () => undefined
                        });
                    """
                }
            )
        except Exception:
            pass

        return driver

    # --------------------------
    # Chromedriver 버전 감지   
    # --------------------------      
    def get_chrome_major_version(self):
        """
        Windows 레지스트리에서 Chrome 버전을 읽고
        메이저 버전(int)을 반환
        예: 147.0.7727.102 -> 147
        """
        reg_paths = [
            (winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon", "version"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Google\Chrome\BLBeacon", "version"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Google\Chrome\BLBeacon", "version"),
        ]

        for root, path, name in reg_paths:
            try:
                with winreg.OpenKey(root, path) as key:
                    version, _ = winreg.QueryValueEx(key, name)
                    if version:
                        major = int(str(version).split(".")[0])
                        self.write_log(f"감지된 Chrome 버전: {version} (major={major})")
                        return major
            except Exception:
                continue

        raise Exception("설치된 Chrome 버전을 찾을 수 없습니다.")

    def get_user_friendly_error(self, step, exception):
        if isinstance(exception, TimeoutException):
            return f"{step} 단계에서 응답 시간이 초과되었습니다."
        elif isinstance(exception, NoSuchElementException):
            return f"{step} 단계에서 필요한 요소를 찾지 못했습니다."
        elif isinstance(exception, ElementClickInterceptedException):
            return f"{step} 단계에서 다른 요소에 가려 클릭할 수 없었습니다."
        elif isinstance(exception, ElementNotInteractableException):
            return f"{step} 단계에서 요소를 조작할 수 없었습니다."
        elif isinstance(exception, StaleElementReferenceException):
            return f"{step} 단계에서 화면 갱신으로 요소 참조가 끊어졌습니다."
        elif isinstance(exception, WebDriverException):
            return f"{step} 단계에서 브라우저 오류가 발생했습니다."
        else:
            return f"{step} 단계에서 오류 발생: {exception}"

    def set_row_result(self, row_idx, status=None, step=None, error=None, **extra):
        if status is not None:
            self.data_frame.at[row_idx, "상태"] = status
        if step is not None:
            self.data_frame.at[row_idx, "작업단계"] = step
        if error is not None:
            self.data_frame.at[row_idx, "에러"] = error

        for key, value in extra.items():
            if key not in self.data_frame.columns:
                self.data_frame[key] = ""
            self.data_frame.at[row_idx, key] = value

    # --------------------------
    # 서비스별 처리
    # --------------------------

    #배달의민족 처리 함수 
    def process_baemin(self, driver, row, row_idx):
        step = "초기화"
        try:
            settings = self.crawl_settings["배달의민족"]
            user_id = str(row.get("ID", "")).strip()
            user_pw = str(row.get("PW", "")).strip()

            self.write_log(f"[행 {row_idx + 1}] 배달의민족 시작")

            # 1. 로그인
            step = "배민 로그인"
            self.baemin_login(driver, user_id, user_pw, row_idx)

            # 2. owner 페이지 수집
            if any([
                settings.get("연락처"),
                settings.get("사업자번호"),
                settings.get("주소")
            ]):
                step = "owner 페이지 수집"
                self.collect_baemin_owner_page(driver, row_idx, settings)

            # 3. shops 페이지 수집
            if any([
                settings.get("가맹점명"),
                settings.get("샵넘버"),
                settings.get("광고서비스사용유무")
            ]):
                step = "shops 페이지 이동"
                self.go_to_baemin_page(
                    driver,
                    "https://self.baemin.com/shops",
                    row_idx,
                    step
                )

                step = "shops 페이지 수집"
                self.collect_baemin_shops_page(driver, row_idx, settings)

            # 4. baro-pay 페이지 수집
            if settings.get("주문접수채널"):
                step = "baro-pay 페이지 이동"
                self.go_to_baemin_page(
                    driver,
                    "https://self.baemin.com/mypage/baro-pay",
                    row_idx,
                    step
                )

                step = "baro-pay 페이지 수집"
                self.collect_baemin_baro_pay_page(driver, row_idx, settings)

            self.set_row_result(row_idx, status="성공", step="완료", error="")
            self.write_log(f"[행 {row_idx + 1}] 배달의민족 성공")

        except Exception as e:
            friendly_error = self.get_user_friendly_error(step, e)
            self.set_row_result(row_idx, status="실패", step=step, error=friendly_error)
            self.write_log(f"[행 {row_idx + 1}] 배달의민족 실패 - {friendly_error}")

    #요기요 처리 함수
    def process_yogiyo(self, driver, row, row_idx):
        step = "초기화"
        try:
            self.write_log(f"[행 {row_idx + 1}] 요기요 시작")
            self.set_row_result(row_idx, status="실패", step="미구현", error="요기요는 아직 구현되지 않았습니다.")
            self.write_log(f"[행 {row_idx + 1}] 요기요 미구현")
        except Exception as e:
            friendly_error = self.get_user_friendly_error(step, e)
            self.set_row_result(row_idx, status="실패", step=step, error=friendly_error)
            self.write_log(f"[행 {row_idx + 1}] 요기요 실패 - {friendly_error}")

    #쿠팡이츠 처리 함수
    def process_coupang(self, driver, row, row_idx):
        step = "초기화"
        try:
            settings = self.crawl_settings["쿠팡이츠"]
            user_id = str(row.get("ID", "")).strip()
            user_pw = str(row.get("PW", "")).strip()

            self.write_log(f"[행 {row_idx + 1}] 쿠팡이츠 시작")

            step = "쿠팡이츠 로그인 페이지 접속"
            self.set_row_result(row_idx, step=step)
            driver.get("https://store.coupangeats.com/merchant/login")

            step = "아이디 입력"
            id_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[1]/input'))
            )
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(3)

            step = "비밀번호 입력"
            pw_input = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/div[2]/input')
            pw_input.clear()
            pw_input.send_keys(user_pw)
            time.sleep(3)

            step = "로그인 버튼 클릭"
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[2]/div/div/div/form/button'))
            )
            login_button.click()
            time.sleep(3)

            login_error = driver.find_elements(By.XPATH, '//*[contains(text(), "아이디 혹은 비밀번호가 일치하지 않습니다.")]')
            if login_error:
                self.set_row_result(row_idx, status="실패", step="로그인 확인", error="로그인 실패")
                return

            step = "가게 목록 페이지 이동"
            self.set_row_result(row_idx, step=step)
            driver.get("https://store.coupangeats.com/merchant/management/stores/")
            time.sleep(3)

            if settings.get("가맹점명") or settings.get("샵넘버") or settings.get("상태"):
                step = "가게 정보 수집"
                self.set_row_result(row_idx, step=step)

                store_list = self.extract_coupang_store_list(driver)

                for idx, store in enumerate(store_list, start=1):
                    if settings.get("가맹점명"):
                        self.set_row_result(row_idx, **{f"가맹점명_{idx}": store.get("가맹점명", "")})

                    if settings.get("샵넘버"):
                        self.set_row_result(row_idx, **{f"샵넘버_{idx}": store.get("샵넘버", "")})

                    if settings.get("상태"):
                        self.set_row_result(row_idx, **{f"상태_{idx}": store.get("상태", "")})

            self.set_row_result(row_idx, status="성공", step="완료", error="")
            self.write_log(f"[행 {row_idx + 1}] 쿠팡이츠 성공")

        except Exception as e:
            friendly_error = self.get_user_friendly_error(step, e)
            self.set_row_result(row_idx, status="실패", step=step, error=friendly_error)
            self.write_log(f"[행 {row_idx + 1}] 쿠팡이츠 실패 - {friendly_error}")

    #땡겨요 처리 함수 
    def process_ddangyo(self, driver, row, row_idx):
        step = "초기화"
        try:
            self.write_log(f"[행 {row_idx + 1}] 땡겨요 시작")
            self.set_row_result(row_idx, status="실패", step="미구현", error="땡겨요는 아직 구현되지 않았습니다.")
            self.write_log(f"[행 {row_idx + 1}] 땡겨요 미구현")
        except Exception as e:
            friendly_error = self.get_user_friendly_error(step, e)
            self.set_row_result(row_idx, status="실패", step=step, error=friendly_error)
            self.write_log(f"[행 {row_idx + 1}] 땡겨요 실패 - {friendly_error}")


    # --------------------------
    # 구글 스프레드시트 오픈 관련 함수 
    # --------------------------

    def open_googlesheet(self):
        try:
            sheet_url = "https://docs.google.com/spreadsheets/d/1iCMikJwc3FqxNtly8zkh6GFdX4luvs155TH2mbsVxd0/edit?gid=0#gid=0"
            webbrowser.open(sheet_url)
            self.write_log("구글 스프레드시트를 브라우저에서 열었습니다.")
        except Exception as e:
            self.write_log(f"구글 스프레드시트 열기 실패: {e}")
            messagebox.showerror("오류", f"구글 스프레드시트를 여는 중 오류가 발생했습니다.\n{e}")



    # --------------------------
    # 배달의 민족 관련 함수 모음 
    # --------------------------

    #배달의민족 로그인 함수 
    def baemin_login(self, driver, user_id, user_pw, row_idx):
        step = "배민 로그인 페이지 접속"
        try:
            self.set_row_result(row_idx, step=step)
            driver.get("https://self.baemin.com/mypage/owner")

            step = "로그인 페이지 안정화"
            self.set_row_result(row_idx, step=step)
            self.stabilize_baemin_page(driver, wait_time=2)

            step = "아이디 입력"
            self.set_row_result(row_idx, step=step)
            id_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@type="text"]'))
            )
            id_input.clear()
            id_input.send_keys(user_id)

            step = "비밀번호 입력"
            self.set_row_result(row_idx, step=step)
            pw_input = driver.find_element(By.XPATH, '//input[@type="password"]')
            pw_input.clear()
            pw_input.send_keys(user_pw)

            step = "로그인 버튼 클릭"
            self.set_row_result(row_idx, step=step)
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//form//button'))
            )
            login_button.click()

            step = "로그인 후 팝업 정리"
            self.set_row_result(row_idx, step=step)
            self.stabilize_baemin_page(driver, wait_time=3)

            step = "로그인 결과 확인"
            self.set_row_result(row_idx, step=step)
            login_error = driver.find_elements(
                By.XPATH,
                '//*[contains(text(), "아이디 또는 비밀번호가 일치하지 않습니다.")]'
            )
            if login_error:
                raise Exception("로그인 실패")

            # 로그인 성공 후 한 번 더 팝업 정리
            self.close_baemin_popups(driver)

        except Exception as e:
            friendly_error = self.get_user_friendly_error(step, e)
            self.set_row_result(row_idx, status="실패", step=step, error=friendly_error)
            raise

    #배달의민족 팝업 닫기 함수 
    def close_baemin_popups(self, driver):
        try:
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
        except Exception:
            pass

    #배달의민족 Owner Page 수집 
    def collect_baemin_owner_page(self, driver, row_idx, settings):
        step = "owner 페이지 분석"
        try:
            current_url = driver.current_url
            if "/mypage/owner" not in current_url:
                step = "owner 페이지 이동"
                self.go_to_baemin_page(
                    driver,
                    "https://self.baemin.com/mypage/owner",
                    row_idx,
                    step
                )

            step = "owner 페이지 팝업 정리"
            self.set_row_result(row_idx, step=step)
            self.close_baemin_popups(driver)

            step = "owner 상세영역 분석"
            self.set_row_result(row_idx, step=step)
            form_group_map = self.extract_form_group_map(driver)

            # 연락처
            if settings.get("연락처"):
                step = "연락처 수집"
                self.set_row_result(row_idx, step=step)
                try:
                    self.close_baemin_popups(driver)
                    group = form_group_map.get("휴대폰")
                    if group:
                        value = group.find_element(By.CSS_SELECTOR, ".inline-values.flex-1").text.strip()
                        self.set_row_result(row_idx, 연락처=value)
                    else:
                        self.set_row_result(row_idx, 연락처="수집 실패: 항목 없음")
                except Exception as sub_e:
                    self.close_baemin_popups(driver)
                    self.set_row_result(row_idx, 연락처=f"수집 실패: {sub_e}")

            # 사업자번호
            if settings.get("사업자번호"):
                step = "사업자번호 수집"
                self.set_row_result(row_idx, step=step)
                try:
                    self.close_baemin_popups(driver)
                    group = form_group_map.get("사업자등록번호")
                    if group:
                        value = group.find_element(By.CSS_SELECTOR, ".inline-values.flex-1").text.strip()
                        self.set_row_result(row_idx, 사업자번호=value)
                    else:
                        self.set_row_result(row_idx, 사업자번호="수집 실패: 항목 없음")
                except Exception as sub_e:
                    self.close_baemin_popups(driver)
                    self.set_row_result(row_idx, 사업자번호=f"수집 실패: {sub_e}")

            # 주소
            if settings.get("주소"):
                step = "주소 수집"
                self.set_row_result(row_idx, step=step)
                try:
                    self.close_baemin_popups(driver)
                    group = form_group_map.get("소재지")
                    if group:
                        dl_element = group.find_element(By.CSS_SELECTOR, "dl.DataList.mt-2.self-ds")
                        address_map = self.parse_dl_data(dl_element)

                        base_address = address_map.get("기본주소", "")
                        jibun_address = address_map.get("지번주소", "")
                        detail_address = address_map.get("상세주소", "")
                        zipcode = address_map.get("우편번호", "")

                        self.set_row_result(
                            row_idx,
                            주소=base_address,
                            지번주소=jibun_address,
                            상세주소=detail_address,
                            우편번호=zipcode
                        )
                    else:
                        self.set_row_result(row_idx, 주소="수집 실패: 항목 없음")
                except Exception as sub_e:
                    self.close_baemin_popups(driver)
                    self.set_row_result(row_idx, 주소=f"수집 실패: {sub_e}")

        except Exception:
            raise
    
    #배달의민족 Shop 페이지 수집 
    def collect_baemin_shops_page(self, driver, row_idx, settings):
        step = "shops 페이지 분석"
        try:
            self.set_row_result(row_idx, step=step)

            self.close_baemin_popups(driver)

            step = "가게 목록 수집"
            self.set_row_result(row_idx, step=step)
            store_list = self.extract_baemin_store_list(driver)

            if not store_list:
                self.write_log(f"[행 {row_idx + 1}] shops 페이지에서 가게 목록을 찾지 못했습니다.")
                return

            for idx, store in enumerate(store_list, start=1):
                self.close_baemin_popups(driver)

                if settings.get("가맹점명"):
                    self.set_row_result(row_idx, **{f"가맹점명_{idx}": store.get("가맹점명", "")})

                if settings.get("샵넘버"):
                    self.set_row_result(row_idx, **{f"샵넘버_{idx}": store.get("샵넘버", "")})

                if settings.get("광고서비스사용유무"):
                    self.set_row_result(row_idx, **{f"광고서비스사용유무_{idx}": store.get("광고서비스사용유무", "")})

        except Exception:
            raise

    #배달의민족 baro-pay 페이지 수집 
    def collect_baemin_baro_pay_page(self, driver, row_idx, settings):
        step = "baro-pay 페이지 분석"
        try:
            current_url = driver.current_url
            if "/mypage/baro-pay" not in current_url:
                step = "baro-pay 페이지 이동"
                self.go_to_baemin_page(
                    driver,
                    "https://self.baemin.com/mypage/baro-pay",
                    row_idx,
                    step
                )

            step = "baro-pay 페이지 팝업 정리"
            self.set_row_result(row_idx, step=step)
            self.close_baemin_popups(driver)

            step = "baro-pay 상세영역 분석"
            self.set_row_result(row_idx, step=step)
            form_group_map = self.extract_form_group_map(driver)

            if settings.get("주문접수채널"):
                step = "주문접수채널 수집"
                self.set_row_result(row_idx, step=step)
                try:
                    self.close_baemin_popups(driver)
                    group = form_group_map.get("주문 접수채널")
                    if group:
                        value = group.find_element(By.CSS_SELECTOR, ".inline-values.flex-1.flex-1").text.strip()
                        self.set_row_result(row_idx, 주문접수채널=value)
                    else:
                        self.set_row_result(row_idx, 주문접수채널="수집 실패: 항목 없음")
                except Exception as sub_e:
                    self.close_baemin_popups(driver)
                    self.set_row_result(row_idx, 주문접수채널=f"수집 실패: {sub_e}")

        except Exception:
            raise

    #배달의민족 팝업 닫기 
    def close_baemin_popups(self, driver):
        try:
            close_buttons = driver.find_elements(By.XPATH, '//button[@aria-label="닫기"]')
            for button in close_buttons:
                try:
                    if button.is_displayed():
                        try:
                            WebDriverWait(driver, 3).until(
                                lambda d: button.is_displayed() and button.is_enabled()
                            )
                            button.click()
                        except Exception:
                            try:
                                driver.execute_script("arguments[0].click();", button)
                            except Exception:
                                pass
                        time.sleep(0.3)
                except Exception:
                    continue
        except Exception:
            pass

    # 배달의민족 페이지 안정화 및 진입 후 팝업 반복 정리
    def stabilize_baemin_page(self, driver, wait_time=1.5):
        time.sleep(wait_time)
        self.close_baemin_popups(driver)
        time.sleep(0.5)
        self.close_baemin_popups(driver)
        time.sleep(0.3)

    #배달의민족 페이지 이동 공통함수 
    def go_to_baemin_page(self, driver, url, row_idx, step):
        self.set_row_result(row_idx, step=step)
        driver.get(url)
        self.stabilize_baemin_page(driver, wait_time=2)

    #배달의민족 form-group 관련 함수 
    def extract_form_group_map(self, driver):
        result = {}
        groups = driver.find_elements(By.CLASS_NAME, "form-group")

        for group in groups:
            try:
                label = group.find_element(By.CLASS_NAME, "form-label").text.strip()
                result[label] = group
            except:
                continue

        return result

    #배달의민족 주소 파싱 함수 
    def parse_dl_data(self, dl_element):
        data = {}
        children = dl_element.find_elements(By.XPATH, "./dt | ./dd")

        current_key = None
        for child in children:
            tag_name = child.tag_name.lower().strip()
            text = child.text.strip()

            if tag_name == "dt":
                current_key = text
            elif tag_name == "dd" and current_key:
                data[current_key] = text
                current_key = None

        return data

    #배달의민족 가맹점명 / 샵넘버 / 광고서비스 조합 함수 
    def extract_baemin_store_list(self, driver):
        self.close_baemin_popups(driver)

        store_list = []

        # 각 가게 카드 단위로 수집
        shop_cards = driver.find_elements(By.CSS_SELECTOR, "div.ShopCard-module__BgLt")

        for card in shop_cards:
            merchant_name = ""
            shop_number = ""
            ad_service = ""

            # 가맹점명
            try:
                merchant_elem = card.find_element(
                    By.CSS_SELECTOR,
                    "p.c_qx9u_13c33de7.Typography_b_r4ax_1bisyd49.Typography_b_r4ax_1bisyd4r.Typography_b_r4ax_1bisyd44j.Typography_b_r4ax_1bisyd44m"
                )
                merchant_name = (merchant_elem.text or merchant_elem.get_attribute("textContent") or "").strip()
            except Exception:
                merchant_name = ""

            # 샵넘버
            try:
                shop_number_elem = card.find_element(
                    By.CSS_SELECTOR,
                    "span.c_qx9u_13c33de7.Typography_b_r4ax_1bisyd49.Typography_b_r4ax_1bisyd4q.Typography_b_r4ax_1bisyd44j"
                )
                raw_shop_number = (shop_number_elem.text or shop_number_elem.get_attribute("textContent") or "").strip()

                # 숫자만 추출하고 싶다면 아래 사용
                match = re.search(r'(\d{5,})', raw_shop_number)
                if match:
                    shop_number = match.group(1)
                else:
                    shop_number = raw_shop_number
            except Exception:
                shop_number = ""

            # 광고서비스 사용 유무
            try:
                ad_link = card.find_element(By.CSS_SELECTOR, "a.ShopCard-module__jp89")
                ad_value_elem = ad_link.find_element(
                    By.CSS_SELECTOR,
                    "span.c_qx9u_13c33de7.Typography_b_r4ax_1bisyd49.Typography_b_r4ax_1bisyd4q.Typography_b_r4ax_1bisyd44j.TextListItem_b_r4ax_n197m76 "
                    "div.Flex_c_qx9u_bbdidai.Flex_c_qx9u_bbdidak.Flex_c_qx9u_bbdida2.TextListItem_b_r4ax_n197m7a"
                )
                ad_service = (ad_value_elem.text or ad_value_elem.get_attribute("textContent") or "").strip()
            except Exception:
                ad_service = ""

            store_list.append({
                "가맹점명": merchant_name,
                "샵넘버": shop_number,
                "광고서비스사용유무": ad_service
            })

        return store_list
    

    # --------------------------
    # 쿠팡이츠 관련 함수 모음 
    # --------------------------

    def extract_coupang_store_list(self, driver):
        store_list = []

        # ul.store-list 대기
        store_ul = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "ul.store-list"))
        )

        li_elements = store_ul.find_elements(By.XPATH, "./li")

        for li in li_elements:
            try:
                # 헤더 행은 a.store-link가 없으므로 제외
                store_link_elements = li.find_elements(By.CSS_SELECTOR, "a.store-link")
                if not store_link_elements:
                    continue

                store_name = ""
                shop_number = ""
                store_status = ""

                # 가맹점명
                try:
                    store_name = store_link_elements[0].text.strip()
                except Exception:
                    store_name = ""

                # span 목록
                span_elements = li.find_elements(By.XPATH, "./span")

                # 첫 번째 span = 샵넘버
                if len(span_elements) >= 1:
                    try:
                        shop_number = span_elements[0].text.strip()
                    except Exception:
                        shop_number = ""

                # 상태
                try:
                    status_elem = li.find_element(By.CSS_SELECTOR, "span.red-font.mw-145")
                    store_status = status_elem.text.strip()
                except Exception:
                    store_status = ""

                store_list.append({
                    "가맹점명": store_name,
                    "샵넘버": shop_number,
                    "상태": store_status
                })

            except Exception:
                continue

        return store_list



if __name__ == "__main__":
    root = tk.Tk()
    root.title("Information Crawling")
    root.geometry("1300x850")
    app = MultiCrawlerApp(root)
    app.pack(fill="both", expand=True)
    root.mainloop()