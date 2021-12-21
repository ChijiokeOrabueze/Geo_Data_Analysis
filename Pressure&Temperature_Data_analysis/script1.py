import xlrd, pandas
from openpyxl import load_workbook


# with open("Book.xlsx", 'rb') as file:
#     content = file.readlines()
#
# content = content[0].decode("utf-8")
# print(str(content))

def save_entries(dict_t):
    df = pandas.DataFrame(dict_t)
    writer = pandas.ExcelWriter("castsQC_boss.xlsx", engine = 'openpyxl')
    writer.book = load_workbook("castsQC_boss.xlsx")
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    reader = pandas.read_excel(r'castsQC_boss.xlsx')
    df.to_excel(writer, index = False, header = False, startrow = len(reader)+1)
    writer.close()

wb = xlrd.open_workbook("casts500_CJ.xlsx")
data = {"S/N":[], "Latitude":[], "Longitude":[], "Top Depth":[], "Top Temperature":[], "Bottom Depth":[],
        "Bottom Temperature":[], "CAST":[], "NODC Cruise ID":[], "Date":[]}
num = 3
for index in range(4,len(wb.sheet_names())):
    sheet = wb.sheet_by_index(index)
    lat = ""
    long = ""
    td = ""
    tt = ""
    bd = ""
    bt = ""
    cast = ""
    nci = ""
    date = ""
    for c in range(sheet.ncols):
        for i in range(sheet.nrows):
            if str(sheet.cell_value(i,c)).strip().lower() == "cast":
                cast = int(sheet.cell_value(i,c+2))
                print(i,c)
                print(int(sheet.cell_value(i,c+2)))
            elif str(sheet.cell_value(i,c)).strip().lower() == "nodc cruise id":
                nci = sheet.cell_value(i,c+2)
                print(sheet.cell_value(i,c+2))
            elif str(sheet.cell_value(i,c)).strip().lower() == "latitude":
                lat = sheet.cell_value(i,c+2)
                print(sheet.cell_value(i,c+2))
            elif str(sheet.cell_value(i,c)).strip().lower() == "longitude":
                long = sheet.cell_value(i,c+2)
                print(sheet.cell_value(i,c+2))
            elif str(sheet.cell_value(i,c)).strip().lower() == "year":
                year = int(sheet.cell_value(i,c+2))
            elif str(sheet.cell_value(i,c)).strip().lower() == "month":
                month = int(sheet.cell_value(i,c+2))
                if len(str(month).strip()) == 1:
                    month = "0" + str(month).strip()
            elif str(sheet.cell_value(i,c)).strip().lower() == "day":
                day = int(sheet.cell_value(i,c+2))
                if len(str(day).strip()) == 1:
                    day = "0" + str(day).strip()
            elif str(sheet.cell_value(i,c)).strip().lower() == "depth":
                top_depth = sheet.cell_value(i+3,c)
                bottom_depth = sheet.cell_value(sheet.nrows-1,c)
                if top_depth.is_integer() == True:
                    top_depth = int(top_depth)
                if bottom_depth.is_integer() == True:
                    bottom_depth = int(bottom_depth)
                td = top_depth
                bd = bottom_depth
            elif str(sheet.cell_value(i,c)).strip().lower() == "temperatur":
                top_temp = sheet.cell_value(i+3,c)
                bottom_temp = sheet.cell_value(sheet.nrows-1,c)
                if top_temp.is_integer() == True:
                    top_temp = int(top_temp)
                if bottom_temp.is_integer() == True:
                    bottom_temp = int(bottom_temp)
                tt = top_temp
                bt = bottom_temp
    date = f'{day}/{month}/{year}'
    data["S/N"].append(num)
    data["Latitude"].append(lat)
    data["Longitude"].append(long)
    data["Top Depth"].append(td)
    data["Top Temperature"].append(tt)
    data["Bottom Depth"].append(bd)
    data["Bottom Temperature"].append(bt)
    data["CAST"].append(cast)
    data["NODC Cruise ID"].append(nci)
    data["Date"].append(date)
    num+=1


save_entries(data)
# with open("sample.txt" , 'w') as file:
#     sheet.dump(f = file)
# sheet = wb.sheet_names()


# print(sheet.cell_value(0,0))
# print(type(sheet.cell_value(0,0)))
#
# print(sheet.nrows)
