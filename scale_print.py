import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32print
import win32api
import fitz  # PyMuPDF
import time

def wait_for_job_completion(printer_name, job_id, timeout=60):
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        elapsed = 0
        while elapsed < timeout:
            jobs = win32print.EnumJobs(hPrinter, 0, -1, 1)
            job_ids = [job['JobId'] for job in jobs]
            if job_id not in job_ids:
                return True
            time.sleep(1)
            elapsed += 1
    finally:
        win32print.ClosePrinter(hPrinter)
    return False

def resize_pdf(input_path, output_path, width, height):
    try:
        src_doc = fitz.open(input_path)
        dst_doc = fitz.open()

        for page in src_doc:
            src_rect = page.rect
            scale = min(width / src_rect.width, height / src_rect.height)
            new_width = src_rect.width * scale
            new_height = src_rect.height * scale
            x_offset = (width - new_width) / 2
            y_offset = (height - new_height) / 2
            matrix = fitz.Matrix(scale, scale)
            dst_page = dst_doc.new_page(width=width, height=height)
            render_rect = fitz.Rect(x_offset, y_offset, x_offset + new_width, y_offset + new_height)
            dst_page.show_pdf_page(render_rect, src_doc, page.number, matrix=matrix, keep_proportion=True)

        dst_doc.save(output_path)
    except Exception as e:
        raise Exception("PDFリサイズ中にエラー発生:", str(e))
    finally:
        if 'src_doc' in locals():
            src_doc.close()
        if 'dst_doc' in locals():
            dst_doc.close()

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDFファイル", "*.pdf")], title="PDFファイルを選択してください")
    if file_path:
        pdf_path.set(file_path)

def select_printer():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    printer_names = [printer[2] for printer in printers]

    def set_printer():
        selected = printer_listbox.curselection()
        if selected:
            printer_name.set(printer_names[selected[0]])
            printer_window.destroy()

    printer_window = tk.Toplevel(root)
    printer_window.title("プリンターを選択")
    printer_listbox = tk.Listbox(printer_window, selectmode=tk.SINGLE)
    printer_listbox.pack(fill=tk.BOTH, expand=True)

    for name in printer_names:
        printer_listbox.insert(tk.END, name)

    select_button = tk.Button(printer_window, text="選択", command=set_printer)
    select_button.pack()

def print_pdf():
    if not pdf_path.get():
        messagebox.showerror("エラー", "PDFファイルを選択してください。")
        return
    if not printer_name.get():
        messagebox.showerror("エラー", "プリンターを選択してください。")
        return

    try:
        paper = paper_option.get()
        paper_sizes = {
            "A1 横": (2384, 1684),
            "A3 横": (1191, 842),
            "A4 横": (842, 595)
        }
        width, height = paper_sizes.get(paper, (842, 595))

        base = os.path.splitext(os.path.basename(pdf_path.get()))[0]
        resized_pdf_path = os.path.join(os.path.dirname(pdf_path.get()), f"{base}_scaled.pdf")

        resize_pdf(pdf_path.get(), resized_pdf_path, width, height)

        hPrinter = win32print.OpenPrinter(printer_name.get())
        try:
            doc_info = ("MyDocument", None, "RAW")
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

        if wait_for_job_completion(printer_name.get(), job_id):
            if os.path.exists(resized_pdf_path):
                os.remove(resized_pdf_path)
        else:
            print("ジョブ完了確認タイムアウト。ファイルは削除されませんでした。")

    except Exception as e:
        messagebox.showerror("エラー", f"印刷に失敗しました: {e}")

def on_drop(event):
    file_path = event.data
    if not os.path.isfile(file_path) or not file_path.lower().endswith('.pdf'):
        messagebox.showerror("エラー", "有効なPDFファイルをドロップしてください。")
        return
    pdf_path.set(file_path)

# メインウィンドウ
root = TkinterDnD.Tk()
root.title("PDFスケーリング印刷")

pdf_path = tk.StringVar()
printer_name = tk.StringVar()
paper_option = tk.StringVar(value="A1 横")

tk.Label(root, text="選択されたPDF:").pack()
tk.Entry(root, textvariable=pdf_path, state="readonly", width=50).pack()
tk.Button(root, text="PDFを選択", command=select_pdf).pack()

tk.Label(root, text="選択されたプリンター:").pack()
tk.Entry(root, textvariable=printer_name, state="readonly", width=50).pack()
tk.Button(root, text="プリンターを選択", command=select_printer).pack()

tk.Label(root, text="印刷用紙サイズ:").pack()
paper_menu = tk.OptionMenu(root, paper_option, "A1 横", "A3 横", "A4 横")
paper_menu.pack()

tk.Button(root, text="印刷", command=print_pdf).pack(pady=10)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

root.mainloop()
