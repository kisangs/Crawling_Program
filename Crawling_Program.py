import tkinter as tk
from tkinter import ttk
from tkinter import font
import Crawling_Request 
import Takedown_Request


def load_frame(frame_class):
    for widget in content_frame.winfo_children():
        widget.destroy()
    frame = frame_class(content_frame)
    frame.pack(fill="both", expand=True)

root = tk.Tk()
root.title("크롤링 프로그램_2025/10/28_Version")
root.geometry("1800x1000")

# 다크 모드 스타일 설정
style = ttk.Style()

# 기본 색상 설정
style.configure("TFrame", background="#2b2b2b")
style.configure("TLabel", background="#2b2b2b", foreground="#ffffff", font=('Helvetica', 14))
style.configure("TButton", background="#3c3c3c", foreground="#000000", font=('Helvetica', 12, 'bold'), justify='center', relief="flat")

# 메인 프레임
main_frame = ttk.Frame(root, style="TFrame")
main_frame.pack(fill="both", expand=True)

# 버튼 프레임
button_frame = ttk.Frame(main_frame, width=200, style="TFrame")
button_frame.pack(side="left", fill="y", padx=10, pady=10)

# 내용 프레임
content_frame = tk.Frame(main_frame, bg="#1e1e1e", bd=2, relief="solid")
content_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

# 다크 모드 버튼들
baemin_crawling_btn = ttk.Button(button_frame, text="정보 크롤링 작업",command=lambda: load_frame(Crawling_Request.Crawling), style="TButton")
baemin_crawling_btn.pack(padx=10, pady=10, fill="x")

baemin_takedown_btn = ttk.Button(button_frame, text="게시중단 요청 작업", command=lambda: load_frame(Takedown_Request.Takedown), style="TButton")
baemin_takedown_btn.pack(padx=10, pady=10, fill="x")

# 폰트 설정 (옵션)
default_font = font.nametofont("TkDefaultFont")
default_font.configure(size=12)

root.option_add("*TButton*Font", default_font)

root.mainloop()