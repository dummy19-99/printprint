import win32api
import win32print
import os

pdf_path = r"C:\Users\woo.jihun\Desktop\Skywalker\sample.pdf"
printer_name = "事務所"

if not os.path.exists(pdf_path):
    print("指定されたPDFが存在しません。")
else:
    confirm = input("印刷しますか？（Y/N）: ").strip().lower()

    if confirm == "y":
        try:
            current_printer = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer_name)
            
            win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)

            win32print.SetDefaultPrinter(current_printer)  # 元に戻す
            print(f"'{printer_name}' に印刷しました。")
        except Exception as e:
            print(f"印刷失敗: {e}")
    else:
        print("印刷はキャンセルされました。")
