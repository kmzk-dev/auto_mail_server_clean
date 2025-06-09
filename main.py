from mail_manager_function import MailManager
import csv
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import tkinter as tk
from tkinter import ttk, messagebox
import json
import sys

class TextRedirector(object):
    """print文の出力をTextウィジェットにリダイレクトするクラス"""
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        # GUIの更新はスレッドセーフであるべきだが、このアプリでは単純化
        self.widget.configure(state='normal')
        self.widget.insert('end', str, (self.tag,))
        self.widget.see('end')  # 自動で最終行までスクロール
        self.widget.configure(state='disabled')

    def flush(self):
        # このメソッドはwriteと対で必要だが、今回は何もしない
        pass

# ===== ▼ここから追加▼ =====
def resource_path(relative_path):
    """ 開発時とPyInstaller実行時の両方でリソースへのパスを解決する """
    try:
        # PyInstallerは一時フォルダを作成し、そのパスを _MEIPASS に格納する
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
# ===== ▲ここまで追加▲ =====


with open(resource_path('server.json'), encoding='utf-8') as f:
    server_data = json.load(f)
accounts_list = list(server_data[0]["ACCOUNT"].keys())
accounts_dict = server_data[0]["ACCOUNT"]
print(accounts_list)
print(accounts_dict)

# global
IMAP_SERVER = server_data[0]["IMAP_SEVER"]
IMAP_PORT = 993
ACCOUNT_NAME = None
PASSWORD = None
mail = None
folders = None

def on_account_selected(event):
    global ACCOUNT_NAME
    global PASSWORD
    ACCOUNT_NAME = combo_account.get()
    PASSWORD = accounts_dict[ACCOUNT_NAME]

def connect_server():
    print("選択中のサーバー")
    print(ACCOUNT_NAME)
    print(PASSWORD)
    global mail
    global folders
    mail = MailManager(IMAP_SERVER,IMAP_PORT,ACCOUNT_NAME,PASSWORD)
    mail.connect()
    folders = mail.GetFolders()
    verify_screen.destroy()

def on_get_folder_list():
    global mail_list
    global directory
    mail_list = []
    directory = combo_folder.get()
    if not directory:
        messagebox.showwarning("警告", "フォルダを選択してください")
        return
    mail.getIds(directory)
    mail_list = mail.fetchHeaders()
    print(mail_list)
        # Treeviewの内容をクリア
    for row in tree.get_children():
        tree.delete(row)
    # mail_list[0]はヘッダーなのでスキップせずカラム名に使う
    for row in mail_list[1:]:
        tree.insert("", "end", values=row)

def on_output_csv():
    if len(mail_list) > 1:
        now_str = (datetime.now() + timedelta(hours=9)).strftime('%Y%m%d_%H%M%S')
        csv_file_path = f'output_{now_str}.csv'
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerows(mail_list)
        messagebox.showinfo("情報", "CSVデータを出力しました")
    else:
        messagebox.showinfo("情報", "directoryにデータがありません")

def on_delete():
    isDelete = messagebox.askyesno("確認", "選択したメールを削除しますか？")
    if isDelete:
        mail.markDeleteFlag()
        mail.Delete()
        on_get_folder_list()

def on_closing_verify():
    verify_screen.destroy()

def on_closing_main():
    mail.disconnect()
    main_screen.destroy()

# =================================================================
#               ▼▼▼ 以下に置き換えるコード ▼▼▼
# =================================================================

# --- ステップ1: 認証ウィンドウの表示 ---
auth_screen = tk.Tk()
auth_screen.title("認証")
auth_screen.geometry("320x160")
# ウィンドウを画面中央に表示するための計算
auth_screen.update_idletasks()
width = auth_screen.winfo_width()
height = auth_screen.winfo_height()
x = (auth_screen.winfo_screenwidth() // 2) - (width // 2)
y = (auth_screen.winfo_screenheight() // 2) - (height // 2)
auth_screen.geometry(f'{width}x{height}+{x}+{y}')
auth_screen.resizable(False, False)

# 既存の関数が 'verify_screen' という名前のグローバル変数を参照するため代入
verify_screen = auth_screen
verify_screen.protocol("WM_DELETE_WINDOW", on_closing_verify)

# UI要素の配置
auth_frame = ttk.Frame(auth_screen, padding=20)
auth_frame.pack(expand=True, fill='both')

ttk.Label(auth_frame, text="アカウントを選択してください:").pack(pady=5)
combo_account = ttk.Combobox(auth_frame, values=accounts_list, state="readonly")
combo_account.pack(fill='x', expand=True, pady=5)
combo_account.bind("<<ComboboxSelected>>", on_account_selected)

style = ttk.Style()
style.configure("Accent.TButton", font=("", 10, "bold"))
ttk.Button(auth_frame, text="接続", command=connect_server, style="Accent.TButton").pack(pady=10)

# 認証ウィンドウのイベントループを開始
# connect_server関数の中で verify_screen.destroy() が呼ばれると、このmainloopは終了する
auth_screen.mainloop()


# --- ステップ2: メインウィンドウの表示 ---
# 認証が成功した場合（folders変数が取得できた場合）のみ、メインの処理に進む
if folders:
    main_screen = tk.Tk()
    main_screen.title("MAIL SERVER MANAGER")
    main_screen.geometry("950x700")  # 高さを少し増やす
    main_screen.minsize(700, 500)
    main_screen.protocol("WM_DELETE_WINDOW", on_closing_main)

    # --- PanedWindowで、メイン画面とログ画面を上下に分割 ---
    paned_window = ttk.PanedWindow(main_screen, orient='vertical')
    paned_window.pack(fill='both', expand=True, padx=10, pady=10)

    # --- 上部ペイン（メインの操作画面） ---
    main_content_frame = ttk.Frame(paned_window)
    paned_window.add(main_content_frame, weight=3) # 操作画面に多くのスペースを割り当てる

    # メイン操作画面のレイアウト設定
    main_content_frame.columnconfigure(0, weight=1)
    main_content_frame.rowconfigure(1, weight=1)

    # --- フォルダ操作フレーム ---
    top_frame = ttk.Labelframe(main_content_frame, text="フォルダ操作", padding=10)
    top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
    top_frame.columnconfigure(1, weight=1)

    ttk.Label(top_frame, text="フォルダ:").grid(row=0, column=0, padx=(0, 5), pady=5)
    combo_folder = ttk.Combobox(top_frame, values=folders, state="readonly")
    combo_folder.grid(row=0, column=1, sticky="ew")
    ttk.Button(top_frame, text="メール取得", command=on_get_folder_list).grid(row=0, column=2, padx=(10, 0))

    # --- メール一覧フレーム ---
    tree_frame = ttk.Labelframe(main_content_frame, text="メール一覧", padding=10)
    tree_frame.grid(row=1, column=0, sticky="nsew", pady=5)
    tree_frame.rowconfigure(0, weight=1)
    tree_frame.columnconfigure(0, weight=1)

    columns = ["date", "fromAdress", "toName", "toAddress", "subject"]
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings")

    tree.heading("date", text="日時")
    tree.column("date", width=140, stretch=False)
    # (他のカラム設定は省略... 前のコードと同じ)
    tree.heading("fromAdress", text="送信元"); tree.column("fromAdress", width=180)
    tree.heading("toName", text="宛先名"); tree.column("toName", width=120)
    tree.heading("toAddress", text="宛先アドレス"); tree.column("toAddress", width=180)
    tree.heading("subject", text="件名"); tree.column("subject", width=300)

    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")

    # --- 操作ボタンフレーム ---
    bottom_frame = ttk.Frame(main_content_frame)
    bottom_frame.grid(row=2, column=0, sticky="e", pady=5)

    ttk.Button(bottom_frame, text="CSV出力", command=on_output_csv).pack(side='left', padx=5)
    style = ttk.Style()
    style.configure("Danger.TButton", foreground="red")
    ttk.Button(bottom_frame, text="表示中のメールを全て削除", command=on_delete, style="Danger.TButton").pack(side='left', padx=5)


    # --- 下部ペイン（ログ出力画面） ---
    log_frame = ttk.Labelframe(paned_window, text="ログ出力", padding=10)
    paned_window.add(log_frame, weight=1) # ログ画面に少しスペースを割り当てる
    log_frame.columnconfigure(0, weight=1)
    log_frame.rowconfigure(0, weight=1)

    # ログ表示用のTextウィジェット
    log_text = tk.Text(log_frame, wrap='word', state='disabled', height=8)
    log_text.grid(row=0, column=0, sticky="nsew")

    log_scrollbar = ttk.Scrollbar(log_frame, orient='vertical', command=log_text.yview)
    log_text.configure(yscrollcommand=log_scrollbar.set)
    log_scrollbar.grid(row=0, column=1, sticky="ns")
    
    # 色分けのためのタグを設定
    log_text.tag_configure("stdout", foreground="#444")
    log_text.tag_configure("stderr", foreground="red")

    # --- 標準出力と標準エラーをリダイレクト ---
    sys.stdout = TextRedirector(log_text, "stdout")
    sys.stderr = TextRedirector(log_text, "stderr")

    # ログエリアに初期メッセージを表示
    print("アプリケーションの準備ができました。")
    print("--------------------------------------------------")

    # --- メインループ開始 ---
    main_screen.mainloop()

# =================================================================
#               ▲▲▲ ここまでを置き換えるコード ▲▲▲
# =================================================================