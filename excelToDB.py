# -*- coding: utf-8 -*
# 插入数据之前做了初步的数据筛选

import xlrd
import pymysql

data = xlrd.open_workbook("2CW-20181111.xlsx")
table_one = data.sheet_by_index(0)  # 根据sheet索引获取sheet的内容
nrows = table_one.nrows

# 连接数据库 并添加cursor游标
db = pymysql.connect(host="localhost", user="root", password="zym2112!", use_unicode=True, charset="utf8")
cursor = db.cursor()

#  创建数据表的语句
sql_createTb = """CREATE TABLE product_info (
                 barcode varchar(50) not null,
                 sellType  varchar(10),
                 wholesalePrice float ,
                 productName varchar(50),
                 referencePurchasePrice float,
                 retailPrice float ,
                 unit varchar(20),
                 specification varchar(20),
                 lowestRetailPrice float,
                 productNature varchar(15),
                 warrantyPeriod int,
                 distributionMethod varchar(10),
                 estimatedDaysOfUse int,
                 grossProfitMargin float )CHARSET=utf8 COLLATE=utf8_bin;
                 """

# cursor.execute("CREATE DATABASE product_info_db;")

cursor.execute("USE product_info_db;")
try:
    cursor.execute("DROP TABLE product_info;")
    cursor.execute(sql_createTb)
except Exception as e:
    db.rollback()
    print(str(e))
    exit(-1)
else:
    print("The original table \"product_info\" has been droped! ")
    print("New table \"product_info\" has been created!")

colName = table_one.row_values(0)

for i in range(1, nrows):

    lineNum = "'" + str(table_one.cell_value(i, 0)) + "'"

    if len(table_one.cell_value(i, 1)) == 13:
        barcode = table_one.cell_value(i, 1)
    else:
        continue

    sellType = "'" + str(table_one.cell_value(i, 2)) + "'"
    wholesalePrice = table_one.cell_value(1, 3)
    if table_one.cell_value(i, 4) != "":
        productName = "'" + str(table_one.cell_value(i, 4)) + "'"
    else:
        productName = "null"
    referencePurchasePrice = table_one.cell_value(i, 5)
    retailPrice = table_one.cell_value(i, 6)
    if float(referencePurchasePrice) > float(retailPrice):
        continue

    if table_one.cell_value(i, 7) != "":
        unit = "'" + str(table_one.cell_value(i, 7)) + "'"
    else:
        unit = "null"
    if table_one.cell_value(i, 8) != "" and table_one.cell_value(i, 8) != "NULL":
        specification = "'" + table_one.cell_value(i, 8) + "'"
    else:
        specification = "null"
    try:
        lowestRetailPrice = float(table_one.cell_value(i, 9))
    except:
        lowestRetailPrice = "null"
    productNature = "'" + table_one.cell_value(i, 10) + "'"

    try:
        if int(table_one.cell_value(i, 11)) == 0:
            warrantyPeriod = "null"
        else:
            warrantyPeriod = "'" + str(table_one.cell_value(i, 11)) + "'"
    except:
        warrantyPeriod = "null"
    distributionMethod = "'" + table_one.cell_value(i, 12) + "'"
    try:
        estimatedDaysOfUse = int(table_one.cell_value(i, 13))
    except:
        estimatedDaysOfUse = "null"
    grossProfitMargin = table_one.cell_value(i, 14)


    sql = "insert into product_info(barcode, sellType, wholesalePrice, productName, referencePurchasePrice,retailPrice,unit,specification,lowestRetailPrice,productNature,warrantyPeriod,distributionMethod, estimatedDaysOfUse,grossProfitMargin) values ({},{},{},{},{},{},{},{},{},{},{},{},{},{});".format(
        barcode, sellType,
        wholesalePrice, productName,
        referencePurchasePrice,
        retailPrice, unit, specification,
        lowestRetailPrice, productNature,
        warrantyPeriod,
        distributionMethod,
        estimatedDaysOfUse,
        grossProfitMargin)

    try:
        cursor.execute(sql)
    except Exception as e:
        db.rollback()
        print(str(e))
        exit(-1)
    else:
        db.commit()  # Transaction submission
        print('Successful insert row ' + str(i) + '! ')

print("============Finished============")
db.close()
cursor.close()
