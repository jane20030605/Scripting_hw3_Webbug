import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox
import re
import requests
import sqlite3
import unicodedata

# 連接 SQLite 資料庫
conn = sqlite3.connect('contacts.db')
cursor = conn.cursor()


def setup_database() -> None:
    """ 初始化資料庫，建立 contacts 資料表。 """
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
    """ 計算字串的顯示寬度，根據東亞寬度判斷。 """
    return sum(2 if unicodedata.east_asian_width(char) in 'WF' else 1 for char in text)


def pad_to_width(text: str, width: int) -> str:
    """ 將字串填充至指定的顯示寬度。 """
    current_width = get_display_width(text)
    padding = width - current_width
    return text + ' ' * padding


def scrape_contacts(url: str) -> str:
    """ 使用 requests 抓取指定 URL 的 HTML 內容。 """
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    return response.text


def parse_contacts(html_content: str) -> list[tuple[str, str, str]]:
    """ 使用正則表達式從 HTML 內容中解析聯絡資訊。 """
    # 姓名、分機、信箱的正則表達式模式
    name_pattern = re.compile(r'<div class="member_name"><a href="[^"]+">([^<]+)</a>')
    ext_pattern = re.compile(r'<div class="member_info_content">([^<]+)</div>')
    email_pattern = re.compile(r'<a href="mailto:([\w.%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})">')

    names = name_pattern.findall(html_content)
    exts = ext_pattern.findall(html_content)
    emails = email_pattern.findall(html_content)

    # 確保每位老師的資料正確對應
    results = []
    for name, ext, email in zip(names, exts, emails):
        results.append((name.strip(), ext.strip(), email.strip()))
    return results


def save_to_database(results: list[tuple[str, str, str]]) -> None:
    """ 將聯絡資訊儲存至 SQLite 資料庫。 """
    for name, ext, email in results:
        cursor.execute("""
            INSERT OR IGNORE INTO contacts (name, ext, email)
            VALUES (?, ?, ?)
        """, (name, ext, email))
    conn.commit()


def display_contacts(results: list[tuple[str, str, str]]) -> None:
    """ 顯示聯絡資訊於 Tkinter 視窗。 """
    output_text.delete("1.0", tk.END)

    headers = ['姓名', '分機', 'Email']
    widths = [20, 20, 30]
    header_line = ''.join(pad_to_width(header, width) for header, width in zip(headers, widths))
    output_text.insert(tk.END, f"{header_line}\n")
    output_text.insert(tk.END, "-" * sum(widths) + "\n")

    for name, ext, email in results:
        row = ''.join(pad_to_width(cell, width) for cell, width in zip([name, ext, email], widths))
        output_text.insert(tk.END, f"{row}\n")


def fetch_data() -> None:
    """ 爬取聯絡資訊並顯示於 Tkinter 視窗，同時儲存至資料庫。 """
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
    """ 關閉視窗時關閉資料庫連線。 """
    cursor.close()
    conn.close()
    root.destroy()

# 主程式介面設置
root = tk.Tk()
root.title("聯絡資訊爬蟲")
root.geometry("640x480")
root.minsize(400, 300)

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=5)
root.grid_columnconfigure(2, weight=1)

url_label = ttk.Label(root, text="URL:")
url_label.grid(row=0, column=0, padx=10, pady=10, sticky="E")

url_var = tk.StringVar(value="https://ai.ncut.edu.tw/p/412-1063-2382.php")
url_entry = ttk.Entry(root, textvariable=url_var)
url_entry.grid(row=0, column=1, padx=10, pady=10, sticky="EW")

fetch_button = ttk.Button(root, text="抓取", command=fetch_data)
fetch_button.grid(row=0, column=2, padx=10, pady=10, sticky="E")

output_text = ScrolledText(root)
output_text.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="NSEW")

root.protocol("WM_DELETE_WINDOW", on_closing)
setup_database()
root.mainloop()
