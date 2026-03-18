import os
import requests
import tkinter as tk
import Crawling_Request
import Takedown_Request
import Crawling_Finance
import Crawling_Finance_Daily
from tkinter import font
from tkinter import ttk, messagebox

# 프레임 생성
def load_frame(frame_class):
    for widget in content_frame.winfo_children():
        widget.destroy()
    frame = frame_class(content_frame)
    frame.pack(fill="both", expand=True)

# Github에서 최신 파일 확인
def get_latest_version_info():
    response = requests.get('https://api.github.com/repos/kisangs/Crawling_Program/releases/latest')
    if response.status_code == 200:
        data = response.json()
        tag_name = data.get('tag_name')
        assets = data.get('assets', [])
        if assets:
            download_url = assets[0].get('browser_download_url')
            return tag_name, download_url
        else:
            return tag_name, None
    else:
        return None, None

# 현재 버전을 파일에서 읽기
def read_current_version():
    try:
        with open('version.txt', 'r') as file:
            return file.read().strip()
    except FileNotFoundError:
        return '1.0.0'

# 현재 버전을 파일에 쓰기
def write_current_version(version):
    with open('version.txt', 'w') as file:
        file.write(version)

current_version = read_current_version()

# Github에서 최신 파일 업데이트 확인 및 다운로드
def check_for_updates():
    latest_version, download_url = get_latest_version_info()
    if latest_version:
        if latest_version > current_version:
            if download_url:
                if tk.messagebox.askyesno("업데이트 확인", f"새 버전 ({latest_version})이 있습니다. 업데이트를 진행하시겠습니까?"):
                    download_new_version(download_url)
                    write_current_version(latest_version)
                    tk.messagebox.showinfo("업데이트 완료", "업데이트가 완료되었습니다. 프로그램을 다시 실행해주세요.")
                    os.startfile('new_version.exe')
                    root.destroy()  # 현재 프로그램 종료
                else:
                    tk.messagebox.showinfo("업데이트 취소", "업데이트가 취소되었습니다.")
            else:
                tk.messagebox.showinfo("업데이트 확인", "새 버전이 있지만, 다운로드 URL을 찾을 수 없습니다.")
        elif latest_version == current_version:
            tk.messagebox.showinfo("업데이트 확인", "현재 최신 버전을 사용 중입니다.")
        else:
            tk.messagebox.showinfo("업데이트 확인", "알 수 없는 버전입니다.")
    else:
        tk.messagebox.showinfo("업데이트 확인", "최신 버전 정보를 가져올 수 없습니다.")

# Github에서 새로운 파일 다운로드
def download_new_version(download_url):
    response = requests.get(download_url, stream=True)
    with open('new_version.exe', 'wb') as file:
        for chunk in response.iter_content(chunk_size=8192):
            file.write(chunk)

root = tk.Tk()
root.title("크롤링 프로그램_2025/10/28_Version")
root.geometry("1800x1000")
style = ttk.Style()
style.configure("TFrame", background="#2b2b2b")
style.configure("TLabel", background="#2b2b2b", foreground="#ffffff", font=('Helvetica', 14))
style.configure("TButton", background="#3c3c3c", foreground="#000000", font=('Helvetica', 12, 'bold'), justify='center', relief="flat")

main_frame = ttk.Frame(root, style="TFrame")
main_frame.pack(fill="both", expand=True)

button_frame = ttk.Frame(main_frame, width=200, style="TFrame")
button_frame.pack(side="left", fill="y", padx=10, pady=10)

content_frame = tk.Frame(main_frame, bg="#1e1e1e", bd=2, relief="solid")
content_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

baemin_crawling_btn = ttk.Button(button_frame, text="정보 크롤링", command=lambda: load_frame(Crawling_Request.Crawling), style="TButton")
baemin_crawling_btn.pack(padx=10, pady=10, fill="x")

baemin_takedown_btn = ttk.Button(button_frame, text="게시중단 요청", command=lambda: load_frame(Takedown_Request.Takedown), style="TButton")
baemin_takedown_btn.pack(padx=10, pady=10, fill="x")

baemin_finance_btn = ttk.Button(button_frame, text="결제금액 확인", command=lambda: load_frame(Crawling_Finance.CrawlingFinance), style="TButton")
baemin_finance_btn.pack(padx=10, pady=10, fill="x")

baemin_finance_Daily_btn = ttk.Button(button_frame, text="일자별 결제금액 확인", command=lambda: load_frame(Crawling_Finance_Daily.CrawlingFinance), style="TButton")
baemin_finance_Daily_btn.pack(padx=10, pady=10, fill="x")

update_check_btn = ttk.Button(button_frame, text="업데이트 확인", command=check_for_updates, style="TButton")
update_check_btn.pack(padx=10, pady=10, fill="x", side="bottom")

default_font = font.nametofont("TkDefaultFont")
default_font.configure(size=12)
root.option_add("*TButton*Font", default_font)

root.mainloop()