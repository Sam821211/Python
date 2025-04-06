# 標準庫
import os #Python用來與作業系統互動，如檔案操作、路徑處理
import sys #提供與Python解譯器互動的功能。ex:sys.argv：讀取指令列參數;sys.exit()：強制結束程式;sys.path：查看模組載入路徑。
from datetime import date, datetime # date處理年月日;datetime：可以包含年月日時分秒，用來產生「現在時間」或格式化。

# 第三方套件
from fpdf import FPDF, XPos, YPos #pip install fpdf2。 FPDF:常用於建立PDF檔案;XPos&YPos:用來控制位置對齊的參考值 
import win32print #pip install pywin32。 Windows專屬模組，通常用於列印功能

# Tkinter GUI
import tkinter as tk #Python內建的GUI套件，用來建立圖形介面視窗
from tkinter import ttk, filedialog, messagebox #ttk:美化外觀的元件;filedialog:提供開啟/儲存檔案的對話框功能;messagebox:彈出提示訊息或警告視窗

def resource_path(relative_path):
    """取得打包後與開發時都適用的資源路徑"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

FONT_PATH = resource_path("msjh.ttc")

class ReceiptApp:
    def __init__(self, root):
        self.root = root
        self.root.title("收據/發票輸入介面 - (新版)")
        self.logo_path = resource_path("logo.png")

        self.printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        self.selected_printer = tk.StringVar(value=win32print.GetDefaultPrinter())

        frame_printer = tk.LabelFrame(root, text="選擇印表機")
        frame_printer.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        tk.Label(frame_printer, text="印表機：").pack(side="left")
        self.printer_menu = ttk.Combobox(frame_printer, textvariable=self.selected_printer, values=self.printers, width=50)
        self.printer_menu.pack(side="left", padx=10)

        frame_company = tk.LabelFrame(root, text="公司資訊")
        frame_company.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        self.company_name = self.create_entry(frame_company, "公司名稱：", 0, 0, 50)
        self.company_address = self.create_entry(frame_company, "地址：", 1, 0, 50)
        self.company_phone = self.create_entry(frame_company, "電話：", 2, 0, 50)
        self.company_owner = self.create_entry(frame_company, "負責人姓名：", 3, 0, 50)

        self.company_name.insert(0, "薩摩亞商動見科技股份有限公司台灣分公司")
        self.company_address.insert(0, "台北市大安區仁愛路四段107號(環球企業大樓14樓) ")
        self.company_phone.insert(0, "0975-221-087")
        self.company_owner.insert(0, "林育賢")

        frame_receipt = tk.LabelFrame(root, text="收據資訊")
        frame_receipt.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.receipt_date = self.create_entry(frame_receipt, "日期：", 0, 0, 30)
        self.receipt_no = self.create_entry(frame_receipt, "收據編號：", 1, 0, 30)
        self.client_no = self.create_entry(frame_receipt, "客戶編號：", 2, 0, 30)
        self.invoice_no = self.create_entry(frame_receipt, "發票號碼：", 3, 0, 30)
        self.payment_date = self.create_entry(frame_receipt, "收款日期：", 4, 0, 30)

        today = date.today().isoformat()
        self.receipt_date.insert(0, today)
        self.payment_date.insert(0, today)
        self.receipt_no.insert(0, "00001")
        self.client_no.insert(0, "51234")
        self.invoice_no.insert(0, "00002")

        frame_client = tk.LabelFrame(root, text="收據人資訊")
        frame_client.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.client_name = self.create_entry(frame_client, "客戶姓名：", 0, 0, 30)
        self.client_address = self.create_entry(frame_client, "地址：", 1, 0, 30)
        self.client_phone = self.create_entry(frame_client, "電話：", 2, 0, 30)

        self.client_name.insert(0, "林岱暘")
        self.client_address.insert(0, "台北市大安區芳蘭路49號9樓")
        self.client_phone.insert(0, "0255790123")

        frame_hospital = tk.LabelFrame(root, text="醫院地址 (如有)")
        frame_hospital.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.hospital_name = self.create_entry(frame_hospital, "醫院名稱：", 0, 0, 30)
        self.hospital_address = self.create_entry(frame_hospital, "地址：", 1, 0, 30)
        self.hospital_phone = self.create_entry(frame_hospital, "電話：", 2, 0, 30)

        self.hospital_name.insert(0, "臺大醫院")
        self.hospital_address.insert(0, "台北市中正區中山南路7號")
        self.hospital_phone.insert(0, "0223123456")

        frame_items = tk.LabelFrame(root, text="項目明細")
        frame_items.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self.tree = ttk.Treeview(frame_items, columns=("desc", "price", "qty", "total"), show="headings")
        for col, text in zip(("desc", "price", "qty", "total"), ("描述", "單價", "數量/小時", "項目合計")):
            self.tree.heading(col, text=text)
            self.tree.column(col, anchor="center")
        self.tree.pack(fill="x")

        frame_input = tk.Frame(root)
        frame_input.grid(row=3, column=0, columnspan=2, pady=5, padx=10, sticky="ew")

        tk.Label(frame_input, text="描述：").grid(row=0, column=0, padx=(0, 5))
        self.entry_desc = tk.Entry(frame_input, width=20)
        self.entry_desc.grid(row=0, column=1, padx=(0, 15))

        tk.Label(frame_input, text="單價：").grid(row=0, column=2, padx=(0, 5))
        self.entry_price = tk.Entry(frame_input, width=10)
        self.entry_price.grid(row=0, column=3, padx=(0, 15))

        tk.Label(frame_input, text="數量/小時：").grid(row=0, column=4, padx=(0, 5))
        self.entry_qty = tk.Entry(frame_input, width=10)
        self.entry_qty.grid(row=0, column=5, padx=(0, 15))

        tk.Button(frame_input, text="新增項目", command=self.add_item).grid(row=0, column=6, padx=(0, 10))
        tk.Button(frame_input, text="刪除項目", command=self.delete_item).grid(row=0, column=7)

        tk.Label(root, text="備註：").grid(row=4, column=0, sticky="w", padx=10)
        self.notes = tk.Text(root, height=3, width=90)
        self.notes.insert("1.0", "這裡可以放一些說明文字，例如付款方式、注意事項等等。")
        self.notes.grid(row=5, column=0, padx=10, sticky="ew")

        frame_total = tk.LabelFrame(root, text="結算")
        frame_total.grid(row=5, column=1, sticky="nsew", padx=10)
        self.subtotal = tk.DoubleVar(value=0.0)
        self.tax_rate = tk.DoubleVar(value=10.0)
        self.tax_amount = tk.DoubleVar(value=0.0)
        self.grand_total = tk.DoubleVar(value=0.0)

        tk.Label(frame_total, text="小計：").grid(row=0, column=0, sticky="e")
        tk.Label(frame_total, textvariable=self.subtotal, width=10, anchor="e").grid(row=0, column=1)

        tk.Label(frame_total, text="(稅) %：").grid(row=1, column=0, sticky="e")
        tax_entry = tk.Entry(frame_total, textvariable=self.tax_rate, width=10, justify="right")
        tax_entry.grid(row=1, column=1)
        tax_entry.bind("<KeyRelease>", lambda e: self.update_total())

        tk.Label(frame_total, text="總計：").grid(row=2, column=0, sticky="e")
        tk.Label(frame_total, textvariable=self.grand_total, width=10, anchor="e").grid(row=2, column=1)

        frame_btns = tk.Frame(root)
        frame_btns.grid(row=6, column=0, columnspan=2, pady=10)
        tk.Button(frame_btns, text="存成 PDF", command=self.generate_pdf).pack(side="left", padx=10)
        tk.Button(frame_btns, text="列印 PDF", command=self.print_pdf).pack(side="left", padx=10)

    def create_entry(self, parent, text, row, col, width):
        tk.Label(parent, text=text).grid(row=row, column=col, sticky="e")
        entry = tk.Entry(parent, width=width)
        entry.grid(row=row, column=col+1, padx=5, pady=2)
        return entry

    def add_item(self):
        try:
            desc = self.entry_desc.get()
            price = float(self.entry_price.get())
            qty = float(self.entry_qty.get())
            total = price * qty
            self.tree.insert("", "end", values=(desc, price, qty, total))
            self.update_total()
        except ValueError:
            messagebox.showerror("錯誤", "請輸入正確數字")

    def delete_item(self):
        for item in self.tree.selection():
            self.tree.delete(item)
        self.update_total()

    def update_total(self):
        try:
            subtotal = sum(float(self.tree.item(child)['values'][3]) for child in self.tree.get_children())
            tax = subtotal * self.tax_rate.get() / 100
            total = subtotal + tax
            self.subtotal.set(round(subtotal, 2))
            self.tax_amount.set(round(tax, 2))
            self.grand_total.set(round(total, 2))
        except tk.TclError:
            pass

    def generate_pdf(self):
        pdf = FPDF()
        pdf.add_page()
        try:
            pdf.add_font('MSJH', '', FONT_PATH)
            pdf.add_font('MSJH', 'B', FONT_PATH)
            pdf.add_font('MSJH', 'I', FONT_PATH)
            pdf.set_font('MSJH', '', 12)
        except:
            messagebox.showerror("錯誤", f"請確認字體檔 {FONT_PATH} 是否存在")
            return

        if os.path.exists(self.logo_path):
            pdf.image(self.logo_path, x=160, y=8, w=40)

        pdf.set_font('MSJH', '', 16)
        pdf.cell(0, 12, "正式收據 Receipt", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(3)

        pdf.set_font('MSJH', '', 12)
        pdf.multi_cell(0, 8, f"公司名稱：{self.company_name.get()}\n負責人：{self.company_owner.get()}\n地址：{self.company_address.get()}\n電話：{self.company_phone.get()}")
        pdf.ln(2)
        pdf.multi_cell(0, 8, f"收據編號：{self.receipt_no.get()}\n發票編號：{self.invoice_no.get()}\n開立日期：{self.receipt_date.get()}\n")
        pdf.multi_cell(0, 8, f"醫院名稱：{self.hospital_name.get()}\n地址：{self.hospital_address.get()}\n電話：{self.hospital_phone.get()}\n")
        pdf.multi_cell(0, 8, f"客戶姓名：{self.client_name.get()}\n地址：{self.client_address.get()}\n電話：{self.client_phone.get()}")

        pdf.ln(5)
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('MSJH', 'B', 12)
        pdf.cell(80, 10, "描述", border=1, align='C', fill=True)
        pdf.cell(30, 10, "單價", border=1, align='C', fill=True)
        pdf.cell(30, 10, "數量", border=1, align='C', fill=True)
        pdf.cell(40, 10, "項目合計", border=1, align='C', fill=True)
        pdf.ln()

        pdf.set_font('MSJH', '', 12)
        for child in self.tree.get_children():
            values = self.tree.item(child)['values']
            pdf.cell(80, 10, str(values[0]), border=1)
            pdf.cell(30, 10, f"{float(values[1]):.2f}", border=1, align='R')
            pdf.cell(30, 10, f"{float(values[2]):.2f}", border=1, align='R')
            pdf.cell(40, 10, f"{float(values[3]):.2f}", border=1, align='R')
            pdf.ln()

        pdf.ln(3)
        pdf.cell(0, 8, f"小計：{self.subtotal.get():.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 8, f"稅率：{self.tax_rate.get():.2f}%", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 8, f"稅額：{self.tax_amount.get():.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 8, f"總計：{self.grand_total.get():.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.ln(3)
        pdf.multi_cell(0, 8, f"備註：\n{self.notes.get('1.0', 'end').strip()}")

        pdf.ln(5)
        pdf.set_font('MSJH', 'I', 11)
        pdf.multi_cell(0, 8, "感謝惠顧！如對本次收據有任何疑問，請與我們聯繫。")
        
        now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        title="存成 PDF",
        initialfile=f"動見收據_{now_str}"
        )

        if not file_path:
            return  # 使用者按取消

        self.pdf_filename = file_path

        pdf.output(self.pdf_filename)
        messagebox.showinfo("成功", f"PDF 已儲存為 {self.pdf_filename}")

    def print_pdf(self):
        self.generate_pdf()
        try:
            printer_name = self.selected_printer.get()
            win32print.SetDefaultPrinter(printer_name)
            os.startfile(self.pdf_filename, 'print')
        except:
            messagebox.showerror("錯誤", "列印時發生問題，請確認系統印表機設定")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReceiptApp(root)
    root.mainloop()
