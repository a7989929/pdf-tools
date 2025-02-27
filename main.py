import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox

def merge_pdfs(odd_pdf_path, even_pdf_path, output_pdf_path):
    try:
        odd_pdf = fitz.open(odd_pdf_path)
        even_pdf = fitz.open(even_pdf_path)
        merged_pdf = fitz.open()

        odd_pages = len(odd_pdf)
        even_pages = len(even_pdf)

        # 確保頁數一致
        if odd_pages != even_pages and odd_pages != even_pages + 1:
            messagebox.showerror("錯誤", "奇數頁與偶數頁的 PDF 頁數不匹配！")
            return

        for i in range(max(odd_pages, even_pages)):
            if i < odd_pages:
                merged_pdf.insert_pdf(odd_pdf, from_page=i, to_page=i)
            if i < even_pages:
                merged_pdf.insert_pdf(even_pdf, from_page=i, to_page=i)

        merged_pdf.save(output_pdf_path)
        merged_pdf.close()
        odd_pdf.close()
        even_pdf.close()
        
        messagebox.showinfo("成功", "PDF 合併完成！")
    
    except Exception as e:
        messagebox.showerror("錯誤", f"合併失敗: {e}")

def select_odd_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        odd_pdf_path.set(file_path)

def select_even_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        even_pdf_path.set(file_path)

def select_output_pdf():
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_pdf_path.set(file_path)

def merge_action():
    if not odd_pdf_path.get() or not even_pdf_path.get() or not output_pdf_path.get():
        messagebox.showwarning("警告", "請選擇所有檔案路徑！")
        return
    
    merge_pdfs(odd_pdf_path.get(), even_pdf_path.get(), output_pdf_path.get())

# 建立 tkinter 視窗
root = tk.Tk()
root.title("PDF 合併工具")
root.geometry("400x300")

# 變數存放選擇的檔案路徑
odd_pdf_path = tk.StringVar()
even_pdf_path = tk.StringVar()
output_pdf_path = tk.StringVar()

# 建立 UI 元件
tk.Label(root, text="選擇奇數頁 PDF:").pack(pady=5)
tk.Entry(root, textvariable=odd_pdf_path, width=40).pack()
tk.Button(root, text="瀏覽", command=select_odd_pdf).pack(pady=5)

tk.Label(root, text="選擇偶數頁 PDF:").pack(pady=5)
tk.Entry(root, textvariable=even_pdf_path, width=40).pack()
tk.Button(root, text="瀏覽", command=select_even_pdf).pack(pady=5)

tk.Label(root, text="輸出合併 PDF:").pack(pady=5)
tk.Entry(root, textvariable=output_pdf_path, width=40).pack()
tk.Button(root, text="選擇存放位置", command=select_output_pdf).pack(pady=5)

tk.Button(root, text="合併 PDF", command=merge_action, bg="lightblue").pack(pady=10)

# 啟動視窗
root.mainloop()
