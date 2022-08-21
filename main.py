import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw

#開啟excel
workbook = xw.Book()

#爬台北、台南、高雄展覽的資料
cities = ["taipei","tainan","kaohsiung"]
for city in cities :
    res = requests.get(f"https://www.funtime.com.tw/localtour/topic/index.php?theme=taiwan-exhibition&target={city}")
    res.encoding="UTF-8"
    html = BeautifulSoup(res.text,"html.parser")
    # print(html)
    titles = html.findAll("span",{"class":"block_title_text"})
    descriptions = html.findAll("div",{"class":"type_c_sub_description"})

    #操作excel
    workbook.sheets["工作表1"].name.replace("工作表1",f"{city}")
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

workbook.sheets['工作表1'].delete()
workbook.save("展覽資訊.xlsx")