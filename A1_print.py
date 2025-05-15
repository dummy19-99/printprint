import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32print
import win32api
import fitz  # PyMuPDF
import time
import win32timezone

def wait_for_job_completion(printer_name, job_id, timeout=60):
    # 指定プリンターのスプール監視
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        elapsed = 0
        while elapsed < timeout:
            jobs = win32print.EnumJobs(hPrinter, 0, -1, 1)
            job_ids = [job['JobId'] for job in jobs]
            if job_id not in job_ids:
                return True  # ジョブが完了した
            time.sleep(1)
            elapsed += 1
    finally:
        win32print.ClosePrinter(hPrinter)
    return False  # タイムアウト

# PDFのサイズを変更する関数（ベクターを維持）
def resize_pdf(input_path, output_path, width, height):
    try:
        src_doc = fitz.open(input_path)
        dst_doc = fitz.open()

        for page in src_doc:
            rect = fitz.Rect(0, 0, width, height)
            dst_page = dst_doc.new_page(width=width, height=height)
            dst_page.show_pdf_page(rect, src_doc, page.number)

        dst_doc.save(output_path)
    except Exception as e:
        raise Exception("エラー", f"PDFのサイズ変更処理中にエラーが発生しました: {e}")
    finally:
        if 'src_doc' in locals():
            src_doc.close()
        if 'dst_doc' in locals():
            dst_doc.close()

# PDFファイルを選択する関数
def select_pdf():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDFファイル", "*.pdf")],  # PDFファイルのみ選択可能
        title="PDFファイルを選択してください"
    )
    if file_path:
        pdf_path.set(file_path)  # 選択したファイルパスを保存

# プリンターを選択する関数
def select_printer():
    # 利用可能なプリンターを取得
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    printer_names = [printer[2] for printer in printers]

    def set_printer():
        selected = printer_listbox.curselection()
        if selected:
            printer_name.set(printer_names[selected[0]])  # 選択したプリンター名を保存
            printer_window.destroy()

    # プリンター選択用のウィンドウを作成
    printer_window = tk.Toplevel(root)
    printer_window.title("プリンターを選択")
    printer_listbox = tk.Listbox(printer_window, selectmode=tk.SINGLE)
    printer_listbox.pack(fill=tk.BOTH, expand=True)

    # プリンター名をリストに追加
    for name in printer_names:
        printer_listbox.insert(tk.END, name)

    # 選択ボタンを作成
    select_button = tk.Button(printer_window, text="選択", command=set_printer)
    select_button.pack()

# PDFを印刷する関数
def print_pdf():
    if not pdf_path.get():
        messagebox.showerror("エラー", "PDFファイルを選択してください。")
        return
    if not printer_name.get():
        messagebox.showerror("エラー", "プリンターを選択してください。")
        return
    if not size_option.get():
        messagebox.showerror("エラー", "印刷サイズを選択してください。")
        return

    try:
        height = int(fixed_height.get())
    except ValueError:
        messagebox.showerror("エラー", "高さは数値で入力してください。")
        return

    size_mapping = {
        "8インチ": (8 * 72, height),
        "10インチ": (10 * 72, height),
        "15インチ": (15 * 72, height)
    }

    try:
        width, height = size_mapping[size_option.get()]
        base = os.path.splitext(os.path.basename(pdf_path.get()))[0]
        resized_pdf_path = os.path.join(os.path.dirname(pdf_path.get()), f"{base}_resized.pdf")

        resize_pdf(pdf_path.get(), resized_pdf_path, width, height)

        # 印刷開始
        hPrinter = win32print.OpenPrinter(printer_name.get())
        try:
            doc_info = ("MyDocumentName", None, "RAW")

            with open(resized_pdf_path, "rb") as f:
                pdf_data = f.read()

            job_id = win32print.StartDocPrinter(hPrinter, 1, doc_info)
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, pdf_data)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)

        finally:
            win32print.ClosePrinter(hPrinter)

        messagebox.showinfo("成功", "印刷を開始しました。")

        # 印刷完了を待って削除
        if wait_for_job_completion(printer_name.get(), job_id):
            if os.path.exists(resized_pdf_path):
                os.remove(resized_pdf_path)
                print(f"印刷完了後にファイル '{resized_pdf_path}' を削除しました。")
        else:
            print("ジョブ完了確認タイムアウト。ファイルは削除されませんでした。")

    except Exception as e:
        messagebox.showerror("エラー", f"印刷に失敗しました: {e}")


# ドラッグ＆ドロップでPDFを選択する関数
def on_drop(event):
    file_path = event.data
    if not os.path.isfile(file_path) or not file_path.lower().endswith('.pdf'):
        messagebox.showerror("エラー", "有効なPDFファイルをドロップしてください。")
        return
    pdf_path.set(file_path)
    
# メインウィンドウの設定
root = TkinterDnD.Tk()
root.title("PDFプリンター")

# グローバル変数の初期化
pdf_path = tk.StringVar()
printer_name = tk.StringVar()
size_option = tk.StringVar()
fixed_height = tk.IntVar(value=842) 


# PDFファイル選択UI

tk.Label(root, text="選択されたPDF:").pack()
tk.Entry(root, textvariable=pdf_path, state="readonly", width=50).pack()
tk.Button(root, text="PDFを選択", command=select_pdf).pack()

# プリンター選択UI
tk.Label(root, text="選択されたプリンター:").pack()
tk.Entry(root, textvariable=printer_name, state="readonly", width=50).pack()
tk.Button(root, text="プリンターを選択", command=select_printer).pack()


# 印刷サイズ選択UI
tk.Label(root, text="印刷サイズを選択:").pack()
size_frame = tk.Frame(root)
size_frame.pack()
sizes = ["8インチ", "10インチ", "15インチ"]
size_option.set("8インチ")  # デフォルト値を設定
for size in sizes:
    tk.Radiobutton(size_frame, text=size, variable=size_option, value=size).pack(anchor=tk.W)

# 固定高さ入力UI
tk.Label(root, text="高さ (ポイント単位):").pack()
tk.Entry(root, textvariable=fixed_height, width=10).pack()

# 印刷ボタン
tk.Button(root, text="印刷", command=print_pdf).pack()

# ドラッグ＆ドロップの設定
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

# メインループの開始
root.mainloop()
