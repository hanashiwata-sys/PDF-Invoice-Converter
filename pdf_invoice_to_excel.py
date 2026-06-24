import glob
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pdfplumber
import re
import pandas as pd

root = tk.Tk()
selected_folder = ""

root.title("PDF→Excel変換ツール")
root.geometry("500x300")

def select_folder():
    
    global selected_folder
    
    folder = filedialog.askdirectory()
    
    if folder:
        
        selected_folder = folder
        
        folder_label.config(
            text=f"選択されたフォルダ:\n{folder}"
        )
        

def execute():
    
    all_data = []
    
    if selected_folder == "":
        
        messagebox.showwarning(
            "警告",
            "フォルダを選択してください"
        )
        
        return
    
    pdf_files = glob.glob(
        os.path.join(selected_folder, "*.pdf")
    )
    
    status_label.config(
        text=f"PDFが{len(pdf_files)}件見つかりました"
    )
    
    if len(pdf_files) == 0:
        
        messagebox.showwarning(
            "警告",
            "PDFがみつかりません"
        )
        
        return
    
    pdf_names = []
    
    for pdf_path in pdf_files:
        
        pdf_names.append(
            os.path.basename(pdf_path)
        )
        
    status_label.config(
        text=(
            f"{len(pdf_files)}件のPDFを処理します\n"
            f"最初のPDF:{os.path.basename(pdf_files[0])}"
        )
    )
    
    for pdf_path in pdf_files:
                
        status_label.config(
            text=f"処理中...\n{os.path.basename(pdf_path)}"
        )

        root.update()
        
        try:
            
            with pdfplumber.open(pdf_path) as pdf:
                  
                for page in pdf.pages:
                    
                    text = page.extract_text()
                    
                    if not text:
                        continue
                        
                    lines = text.split("\n")
                    
                    for line in lines:
                        print(line)
                        
                    if len(lines) > 1:
                        
                        company_name = lines[1]
                        
                    current_date = ""
                    
                    for line in lines:
                        
                        date_match = re.search(
                            r"(\d+月\s+\d+\s+日)",
                            line
                        )
                        
                        if date_match:
                            
                            current_date = date_match.group(1)
                            
                    for line in lines:
                        
                        parts = line.split()
                        
                        if len(parts) < 5:
                            continue
                            
                        if (
                            parts[-1].isdigit()
                            and parts[-2].isdigit()
                            and parts[-4].isdigit()
                        ):
                            
                            amount = int(parts[-1])
                            price = int(parts[-2])
                            quantity = int(parts[-4])
                            
                            unit = parts[-3]
                            
                            product_name = " ".join(parts[:-4])
                            
                            data = {
                                "PDF名": os.path.basename(pdf_path),
                                "会社名": company_name,
                                "日付": current_date,
                                "商品名": product_name,
                                "数量": quantity,
                                "単位": unit,
                                "単価": price,
                                "金額": amount
                            }
                            
                            all_data.append(data)

                
        except Exception as e:
            
            print(e)
            
    df = pd.DataFrame(all_data)
    
    output_file = os.path.join(
        selected_folder,
        "変換結果.xlsx"
    )
    
    df.to_excel(
        output_file,
        index=False
    )
        
    status_label.config(
        text="処理完了！"
    )

    messagebox.showinfo(
        "完了",
        f"{len(all_data)}件を保存しました\n\n{output_file}"
    )


title_label = tk.Label(
    root,
    text = "PDF→excel変換ツール"
)
title_label.pack(pady=20)


select_button = tk.Button(
    root,
    text="フォルダ選択",
    command=select_folder
)
select_button.pack(pady=10)


folder_label = tk.Label(
    root,
    text="フォルダが選択されていません"
)
folder_label.pack(pady=20)


execute_button = tk.Button(
    root,
    text="実行",
    command=execute
)
execute_button.pack(pady=10)


status_label = tk.Label(
    root,
    text="待機中"
)
status_label.pack(pady=20)

root.mainloop()