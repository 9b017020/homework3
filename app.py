import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox
import re
import requests
import unicodedata
import sqlite3

# 連接資
conn = sqlite3.connect('contacts.db')
cursor = conn.cursor()

def setup_database() -> None:
    """
    初始化資料庫，建立資料表 contacts（如果尚未存在）。
    """
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS contacts (
            編號 INTEGER PRIMARY KEY AUTOINCREMENT,
            姓名 TEXT NOT NULL,
            職稱 TEXT NOT NULL,
            Email TEXT NOT NULL UNIQUE
        );
    """)
    conn.commit()

def get_display_width(text: str) -> int:
    """
    計算字串的顯示寬度，以東亞寬度為準。
    """
    return sum(2 if unicodedata.east_asian_width(char) in 'WF' else 1 for char in text)

def pad_to_width(text: str, width: int) -> str:
    """
    將字串填充至指定的顯示寬度。
    """
    current_width = get_display_width(text)
    padding = width - current_width
    return text + ' ' * padding

def scrape_contacts(url: str) -> str:
    """
    從指定的 URL 爬取 HTML 內容。
    """
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    return response.text

def parse_contacts(html_content: str) -> list[tuple[str, str, str]]:
    """
    從 HTML 內容中解析聯絡資訊。
    """
    name_pattern = re.compile(r'<div class="member_name"><a href="[^"]+">([^<]+)</a>')
    title_pattern = re.compile(r'<div class="member_info_content">\s*(.*教授.*?)\s*</div>')
    email_pattern = re.compile(r'<a href="mailto:([\w.%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})">')

    names = name_pattern.findall(html_content)
    titles = [title.strip() for title in title_pattern.findall(html_content) if title.strip()]
    emails = email_pattern.findall(html_content)

    # 確保數據一致性
    results = list(zip(names, titles, emails))
    return results

def save_to_database(results: list[tuple[str, str, str]]) -> None:
    """
    將聯絡資訊儲存到資料庫。
    """
    for name, title, email in results:
        cursor.execute("""
            INSERT OR IGNORE INTO contacts (姓名, 職稱, Email)
            VALUES (?, ?, ?)
        """, (name, title, email))
    conn.commit()

def display_contacts(results: list[tuple[str, str, str]]) -> None:
    """
    在介面中顯示聯絡資訊。
    """
    output_text.delete("1.0", tk.END)

    headers = ['姓名', '職稱', 'Email']
    widths = [20, 30, 28]
    header_line = ''.join(pad_to_width(header, width) for header, width in zip(headers, widths))
    output_text.insert(tk.END, f"{header_line}\n")
    output_text.insert(tk.END, "-" * sum(widths) + "\n")

    for name, title, email in results:
        row = ''.join(pad_to_width(cell, width) for cell, width in zip([name, title, email], widths))
        output_text.insert(tk.END, f"{row}\n")

def fetch_data() -> None:
    """
    抓取聯絡資訊並顯示於介面，儲存至資料庫。
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
    關閉視窗時關閉資料庫連線。
    """
    cursor.close()
    conn.close()
    root.destroy()

# 主程式介面
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

url_var = tk.StringVar(value="https://csie.ncut.edu.tw/content.php?key=86OP82WJQO")
url_entry = ttk.Entry(root, textvariable=url_var)
url_entry.grid(row=0, column=1, padx=10, pady=10, sticky="EW")

fetch_button = ttk.Button(root, text="抓取", command=fetch_data)
fetch_button.grid(row=0, column=2, padx=10, pady=10, sticky="E")

output_text = ScrolledText(root)
output_text.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="NSEW")

root.protocol("WM_DELETE_WINDOW", on_closing)
setup_database()
root.mainloop()
