import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

def resize_pdf():
    path = file_path.get()
    size_choice = size_combo.get()

    if not os.path.exists(path):
        messagebox.showerror("エラー", "PDFファイルが存在しません。")
        return

    # 幅（pt）指定（1インチ = 72pt）
    size_map = {
        "6インチ": 6 * 32,  # 432pt
        "7インチ": 7 * 32,  # 504pt
        "8インチ": 8 * 32   # 576pt
    }
    new_width = size_map.get(size_choice, 576)

    try:
        # PDFファイルを開く
        src_doc = fitz.open(path)
        
        # 新しいPDFを作成
        new_doc = fitz.open()

        for page_num in range(len(src_doc)):
            page = src_doc.load_page(page_num)
            rect = page.rect
            aspect = rect.height / rect.width
            new_height = new_width * aspect

            # 新しいページを追加してサイズを変更
            new_page = new_doc.new_page(width=new_width, height=new_height)
            new_page.show_pdf_page(fitz.Rect(0, 0, new_width, new_height), src_doc, page_num)

        # 保存先を決定
        save_path = os.path.splitext(path)[0] + "_resized.pdf"
        new_doc.save(save_path)

        messagebox.showinfo("保存完了", f"PDFとして保存しました：\n{save_path}")
    except Exception as e:
        messagebox.showerror("エラー", str(e))

# GUI部分
root = tk.Tk()
root.title("PDFサイズ調整保存ツール（PDF出力）")

file_path = tk.StringVar()

tk.Label(root, text="PDFファイルを選択:").pack(pady=5)
tk.Entry(root, textvariable=file_path, width=50).pack()
tk.Button(root, text="ファイル選択", command=lambda: file_path.set(
    filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")]))).pack(pady=5)

tk.Label(root, text="出力幅（インチ）:").pack()
size_combo = ttk.Combobox(root, values=["6インチ", "7インチ", "8インチ"])
size_combo.current(0)
size_combo.pack(pady=5)

tk.Button(root, text="PDFとして保存", command=resize_pdf).pack(pady=10)

root.mainloop()
