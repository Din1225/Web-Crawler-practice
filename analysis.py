import pandas as pd
import xlwings as xw

wb = xw.Book("展覽資訊.xlsx")
cities = ["taipei","tainan","kaohsiung"]
for city in cities :
    df = pd.read_excel("展覽資訊.xlsx",sheet_name=f"{city}")
    df_id = df.set_index("展覽名稱")

    df_id["展覽時間"] = df_id["展覽時間"].str.replace("\n"," ")
    df_id["展覽地點"] = df_id["展覽地點"].str.replace("\n"," ")
    df_id["票價資訊"] = df_id["票價資訊"].str.replace("_x000D_\n\t\t\t\t"," ")

    df_bysite = df_id.sort_values(by=["展覽地點"])
    # print(df_bysite)

    sht = wb.sheets[f"{city}"]
    sht.range("A1").value = df_bysite
    sht.range("A:D").autofit()




# len_total = len(df_bysite["展覽時間"])
# for i in range(0,len_total) :
#     len_ticket = len(df_bysite["票價資訊"][i])
#     for count in range(0,len_ticket) :
#         if (df_bysite["票價資訊"][i][count]=="全") and (df_bysite["票價資訊"][i][count+1]=="票") :
#             df_bysite["全票價"] = df_bysite["票價資訊"][i][count+3] + df_bysite["票價資訊"][i][count+4] + df_bysite["票價資訊"][i][count+5]
#             break



# if (df_bysite["票價資訊"][1][0]=="全") and (df_bysite["票價資訊"][1][1]=="票") :
#             df_bysite["全票價"][1] = df_bysite["票價資訊"][1][3] + df_bysite["票價資訊"][1][4] + df_bysite["票價資訊"][1][5]
# print(df_bysite)
# print(df_bysite["全票價"])  
# print(df_bysite["票價資訊"])    
# print(df_bysite["票價資訊"][1][1])