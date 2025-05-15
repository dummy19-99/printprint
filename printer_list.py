# printer_list.py
import win32print

def main():
    # 接続されているすべてのプリンターを取得
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    print("=== 接続されているプリンター一覧 ===")
    for printer in printers:
        print(f"- {printer[2]}")

    # defaultのプリンターを取得
    default_printer = win32print.GetDefaultPrinter()
    print("\n=== 既定のプリンター ===")
    print(f"{default_printer}")

if __name__ == "__main__":
    main()