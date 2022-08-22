import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw
import tkinter as tk
import tkinter.ttk as ttk
import time

#照地點排序好，最後存檔
def analyze():
    global city
    wb = xw.Book(f"{city}展覽資訊.xlsx")
    df = pd.read_excel(f"{city}展覽資訊.xlsx",sheet_name=f"{city}")
    df_id = df.set_index("展覽名稱")

    df_id["展覽時間"] = df_id["展覽時間"].str.replace("\n"," ")
    df_id["展覽地點"] = df_id["展覽地點"].str.replace("\n"," ")
    df_id["票價資訊"] = df_id["票價資訊"].str.replace("_x000D_\n\t\t\t\t"," ")

    df_bysite = df_id.sort_values(by=["展覽地點"])
    # print(df_bysite)
    sht = wb.sheets[f"{city}"]
    sht.range("A1").value = df_bysite
    sht.range("A:D").autofit()
    wb.save(f"{city}展覽資訊.xlsx")
    #wb.close()

#讓使用者選擇爬台北、台南或高雄展覽的資料
def choice():
    return mycombobox.get()

#按下按鈕後開始爬+整理資料
def button_event():
    global city
    city = choice()
    time.sleep(1)
    find_data()
    analyze()

def find_data():
    #選好後，開始爬資料
    global city
    res = requests.get(f"https://www.funtime.com.tw/localtour/topic/index.php?theme=taiwan-exhibition&target={city}")
    res.encoding="UTF-8"
    html = BeautifulSoup(res.text,"html.parser")
    # print(html)
    titles = html.findAll("span",{"class":"block_title_text"})
    descriptions = html.findAll("div",{"class":"type_c_sub_description"})

    #操作excel
    workbook = xw.Book()
    #workbook.sheets["工作表1"].name.replace("工作表1",f"{city}")
    sheet = workbook.sheets.add(name=f'{city}', before='工作表1')
    sheet.range("A1:D1").value = ["展覽名稱","展覽時間","展覽地點","票價資訊"]

    #寫入展名
    row = 2
    for title in titles :
        sheet.range(f"A{row}").value = title.text.replace("常見問題","")
        row+=1

    #(debug)
    #print(descriptions[1].text[103])
    #print(len(descriptions[1].text))

    #寫入展覽的詳細資料
    #temp1紀錄展覽票價的index temp2紀錄展覽地點的index
    line = 0
    for description in descriptions:
        length = len(description.text)
        for cnt in range(0,length) :
            if ((descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "票") and (descriptions[line].text[cnt+3] == "價")) or ((descriptions[line].text[cnt] == "門") and (descriptions[line].text[cnt+1] == "票") and (descriptions[line].text[cnt+2] == "資") and (descriptions[line].text[cnt+3] == "訊")) or ((descriptions[line].text[cnt] == "單") and (descriptions[line].text[cnt+1] == "展") and (descriptions[line].text[cnt+2] == "票") and (descriptions[line].text[cnt+3] == "價")):
                sheet.range(f"D{line+2}").value = descriptions[line].text[cnt+5:].replace("■","/")
                temp1 = cnt-2
                break
            if cnt == length-1 :
                sheet.range(f"D{line+2}").value = "待公布"

        for cnt in range(0,length) :        
            if (descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "地") and (descriptions[line].text[cnt+3] == "點") :
                sheet.range(f"C{line+2}").value = descriptions[line].text[cnt+5:temp1]
                temp2 = cnt-2
                break
            if cnt == length-1 :
                sheet.range(f"C{line+2}").value = "待公布"

        for cnt in range(0,length) :        
            if ((descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "時") and (descriptions[line].text[cnt+3] == "間")) or ((descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "日") and (descriptions[line].text[cnt+3] == "期")):
                sheet.range(f"B{line+2}").value = descriptions[line].text[cnt+5:temp2]
                break
            if cnt == length-1 :
                sheet.range(f"B{line+2}").value = "待公布"
        line+=1

    #整理格式
    sheet.range("A:D").column_width = 23
    sheet.range("A:D").autofit()
    workbook.save(f"{city}展覽資訊.xlsx")
    #workbook.close()
    
#使用者介面
root = tk.Tk()
root.title("Search-Exhibition")
root.geometry("250x150")

mylabel = tk.Label(root, text='請選擇你要查詢的城市')
mylabel.place(x = 125,y = 30,anchor="center")

mycombobox = ttk.Combobox(root, state="readonly",values=["taipei","tainan","kaohsiung"])
mycombobox.pack(pady=50)
mycombobox.current(0)

mybutton = tk.Button(root, text='確定',command=button_event)
mybutton.place(x = 125,y = 90,anchor="center")

root.mainloop()
