import json
import xlrd
import os

list = []
def start():
    table_name = input("请输入需要导出的表名(无需加.xlsx后缀)：")
    path = os.path.abspath('.') + "\\" + table_name + ".xlsx"   #返回path在当前系统中的绝对路径
    if not os.path.exists(path):
        print("您输入的文件名不存在，请重新输入")
        start()
        return

    table = xlrd.open_workbook(path)    #文件名以及路径，如果路径或者文件名有中文给前面加一个r表示原生字符。
    sheet = table.sheet_by_index(0)     #通过索引顺序获取
    name = sheet.row_values(0)[1]       #.row_values返回由该行中所有单元格的数据组成的列表 这里是表头
    fields = sheet.row_values(1)        # 字段行

    count = 2
    list.append(getOtherXml(sheet, table, fields))

    j = json.dumps(list, sort_keys=False, indent=4, ensure_ascii=False)

    if not os.path.exists(os.path.abspath('.') + "\\output\\"):
        os.makedirs(os.path.abspath('.') + "\\output\\")

    document = open(".//output//" + name + ".json", "w+",  encoding='utf-8')
    document.write(str(j))
    document.close()

    print("已导出文件：" + name + ".json")
    print()

def getOthersheet(sheet, charId):
    fields = sheet.row_values(1)
    count = 2
    list = []
    while count < sheet.col_values(0).__len__():
        dic = {}
        id = sheet.col_values(1)[count]
        if int(id) == charId:
            row_data = sheet.row_values(count)
            for i in range(sheet.ncols):
                dic[fields[i]] = row_data[i]

        if dic.__len__() > 0:
            list.append(dic)
        count = count + 1
    return list

def getOtherXml(otherSheet, otherTable, fields):
    list = []
    count = 2
    while count < otherSheet.col_values(0).__len__():  # 循环遍历
        dic = {}
        row_data = otherSheet.row_values(count)  # 每一行的数据
        for i in range(otherSheet.ncols):  # .ncols获取列表的有效列数
            data = row_data[i]
            if "connet_file" in fields[i]:
                table_other = getConnectXml(data)
                if not table_other:
                    continue

                sheet_other = table_other.sheet_by_index(0)  # 通过索引顺序获取
                connectFieldName = sheet_other.row_values(0)[1]  # .row_values返回由该行中所有单元格的数据组成的列表 这里是表头
                sheet_other = table_other.sheet_by_index(0)
                fields_other = sheet_other.row_values(1)  # 字段行
                _id = sheet_other.col_values(0)[count]
                dic[connectFieldName] = getOtherXml(sheet_other, table_other, fields_other)
                print("已导入关联表格" + data + ".xlsx的数据到json")
            elif "connet" in fields[i]:
                connectFieldName = getConnectSheetName(otherTable, data)
                if not connectFieldName:
                    continue

                sheet_other = otherTable.sheet_by_index(int(data) - 1)
                _id = otherTable.col_values(0)[count]
                dic[connectFieldName] = getOthersheet(sheet_other, int(_id))
                print("已导入关联Sheet" + str(int(data)) + "的数据到json")
            else:
                if isinstance(data, (int, float)):
                    if data == int(data):
                        dic[fields[i]] = int(data)
                    else:
                        dic[fields[i]] = data
                else:
                    dic[fields[i]] = data

        list.append(dic)
        count = count + 1
    return list

#返回关联表格的名字
def getConnectSheetName(data, sheet_index):
    if type(sheet_index) != "Int":
        return

    if len(data.sheets()) < int(sheet_index) - 1:
        print("试图关联的sheet索引" + str(sheet_index) + "非法")
        return None

    sheet = data.sheet_by_index(int(sheet_index) - 1)
    name = sheet.row_values(0)[1]
    return name

#返回关联表格整个table
def getConnectXml(XmlName):
    path = os.path.abspath('.') + "\\" + XmlName + ".xlsx"
    if not os.path.exists(path):
        print("试图关联的表格" + XmlName + ".xlsx不存在")
        return

    table = xlrd.open_workbook(path)
    return table


if __name__ == '__main__':
    while True:
        start()
