import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox
import re
import requests
import sqlite3
import unicodedata

# 連接資料庫
conn = sqlite3.connect('contacts.db')
cursor = conn.cursor()


def setup_database() -> None:
    """
    初始化資料庫，若 contacts 資料表不存在，則創建該資料表。
    """
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            ext TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE
        );
    """)
    conn.commit()


def get_display_width(text: str) -> int:
    """
    計算字串的顯示寬度，東亞字符佔用兩個字符寬度，其他字符佔用一個字符寬度。
    """
    return sum(2 if unicodedata.east_asian_width(char) in 'WF' else 1 for char in text)


def pad_to_width(text: str, width: int) -> str:
    """
    將字串填充至指定寬度，確保每個欄位對齊。
    """
    current_width = get_display_width(text)
    padding = width - current_width
    return text + ' ' * padding


def scrape_contacts(url: str) -> str:
    """
    透過 requests 模組發送 GET 請求，從指定 URL 抓取 HTML 內容。
    """
    response = requests.get(url, timeout=10)
    response.raise_for_status()  # 如果回應不是 200，會引發錯誤
    return response.text


def parse_contacts(html_content: str) -> list[tuple[str, str, str]]:
    """
    使用正規表達式從 HTML 內容中提取聯絡資訊（姓名、分機、Email）。
    """
    name_pattern = re.compile(r'<td class="views-field views-field-title">([^<]+)</td>')
    ext_pattern = re.compile(r'<td class="views-field views-field-field-ext">([^<]+)</td>')
    email_pattern = re.compile(r'<a href="mailto:([^"]+)">')

    names = name_pattern.findall(html_content)
    exts = ext_pattern.findall(html_content)
    emails = email_pattern.findall(html_content)

    # 確保提取的資料是對應的
    results = list(zip(names, exts, emails))
    return results


def save_to_database(results: list[tuple[str, str, str]]) -> None:
    """
    儲存聯絡資訊至 SQLite 資料庫中，避免重複資料。
    """
    for name, ext, email in results:
        cursor.execute("""
            INSERT OR IGNORE INTO contacts (name, ext, email)
            VALUES (?, ?, ?)
        """, (name, ext, email))
    conn.commit()


def display_contacts(results: list[tuple[str, str, str]]) -> None:
    """
    在 Tkinter 視窗中顯示聯絡資訊，資料行對齊。
    """
    output_text.delete("1.0", tk.END)  # 清空目前顯示的內容

    headers = ['姓名', '分機', 'Email']
    widths = [20, 15, 30]
    header_line = ''.join(pad_to_width(header, width) for header, width in zip(headers, widths))
    output_text.insert(tk.END, f"{header_line}\n")
    output_text.insert(tk.END, "-" * sum(widths) + "\n")

    for name, ext, email in results:
        row = ''.join(pad_to_width(cell, width) for cell, width in zip([name, ext, email], widths))
        output_text.insert(tk.END, f"{row}\n")


def fetch_data() -> None:
    """
    處理抓取資料的邏輯，抓取指定 URL 的聯絡資訊並顯示在介面上。
    """
    url = url_var.get()
    if not url:
        messagebox.showwarning("警告", "請輸入 URL！")
        return

    if 'http' not in url or '://' not in url:
        messagebox.showerror("錯誤", "網址格式不正確！")
        return

    try:
        html_content = scrape_contacts(url)
        results = parse_contacts(html_content)
        display_contacts(results)
        save_to_database(results)
    except requests.exceptions.RequestException as e:
        messagebox.showerror("錯誤", f"無法抓取資料：\n{e}")


def on_closing() -> None:
    """
    當關閉視窗時，關閉資料庫連接。
    """
    cursor.close()
    conn.close()
    root.destroy()


# 主程式介面
root = tk.Tk()
root.title("聯絡資訊爬蟲")
root.geometry("640x480")
root.minsize(400, 300)

# 設定欄位的權重
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=5)
root.grid_columnconfigure(2, weight=1)

# URL 輸入框
url_label = ttk.Label(root, text="URL:")
url_label.grid(row=0, column=0, padx=10, pady=10, sticky="E")

url_var = tk.StringVar(value="https://ai.ncut.edu.tw/p/412-1063-2382.php")
url_entry = ttk.Entry(root, textvariable=url_var)
url_entry.grid(row=0, column=1, padx=10, pady=10, sticky="EW")

# 抓取按鈕
fetch_button = ttk.Button(root, text="抓取", command=fetch_data)
fetch_button.grid(row=0, column=2, padx=10, pady=10, sticky="E")

# 顯示聯絡資訊的滾動文字框
output_text = ScrolledText(root)
output_text.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="NSEW")

# 當視窗關閉時執行 on_closing 函數
root.protocol("WM_DELETE_WINDOW", on_closing)

# 初始化資料庫
setup_database()

# 開始主循環
root.mainloop()
