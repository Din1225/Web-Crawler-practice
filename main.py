import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw



res = requests.get("https://www.funtime.com.tw/localtour/topic/index.php?theme=taiwan-exhibition&target=taipei")
res.encoding="UTF-8"
html = BeautifulSoup(res.text,"html.parser")
# print(html)
titles = html.findAll("span",{"class":"block_title_text"})
descriptions = html.findAll("div",{"class":"type_c_sub_description"})


workbook = xw.Book()
sheet = workbook.sheets["工作表1"]

sheet.range("A1:D1").value = ["展覽名稱","展覽時間","展覽地點","票價資訊"]

#寫入展名
row = 2
for title in titles :
    sheet.range(f"A{row}").value = title.text.replace("常見問題","")
    row+=1


# row = 1
# for description in descriptions:
#     sheet.range(f"B{row}").value = description.text
#     row+=2
# sheet.range("A:B").autofit()



#debug
#print(descriptions[1].text[103])
#print(len(descriptions[1].text))


#寫入展覽的詳細資料
#temp1紀錄展覽票價的index temp2紀錄展覽地點的index
line = 0
for description in descriptions:
    length = len(description.text)
    for cnt in range(0,length) :
        if ((descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "票") and (descriptions[line].text[cnt+3] == "價")) or ((descriptions[line].text[cnt] == "門") and (descriptions[line].text[cnt+1] == "票") and (descriptions[line].text[cnt+2] == "資") and (descriptions[line].text[cnt+3] == "訊")):
            sheet.range(f"D{line+2}").value = descriptions[line].text[cnt-2:]
            temp1 = cnt-2
            break
        if cnt == length-1 :
            sheet.range(f"D{line+2}").value = "■ 展覽票價:unknown"

    for cnt in range(0,length) :        
        if (descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "地") and (descriptions[line].text[cnt+3] == "點") :
            sheet.range(f"C{line+2}").value = descriptions[line].text[cnt-2:temp1]
            temp2 = cnt-2
            break
        if cnt == length-1 :
            sheet.range(f"C{line+2}").value = "■ 展覽地點:unknown"

    for cnt in range(0,length) :        
        if ((descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "時") and (descriptions[line].text[cnt+3] == "間")) or ((descriptions[line].text[cnt] == "展") and (descriptions[line].text[cnt+1] == "覽") and (descriptions[line].text[cnt+2] == "日") and (descriptions[line].text[cnt+3] == "期")):
            sheet.range(f"B{line+2}").value = descriptions[line].text[cnt-2:temp2]
            break
        if cnt == length-1 :
            sheet.range(f"B{line+2}").value = "■ 展覽時間:unknown"
    line+=1

#統一欄寬
sheet.range("A1:").autofit()
