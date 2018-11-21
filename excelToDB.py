# -*- coding: utf-8 -*


import xlrd
import pymysql
from datetime import datetime
from xlrd import xldate_as_tuple

data = xlrd.open_workbook("2CW-20181111.xlsx")
table_one = data.sheet_by_index(0)  # 根据sheet索引获取sheet的内容
nrows = table_one.nrows

# 连接数据库 并添加cursor游标
db = pymysql.connect(host="localhost", user="root", password="zym2112!", use_unicode=True, charset="utf8")
cursor = db.cursor()

#  创建数据表的语句
sql_createTb = """CREATE TABLE product_info (
                 barcode varchar(50),
                 sellType  varchar(10),
                 wholesalePrice float ,
                 productName varchar(50),
                 referencePurchasePrice float,
                 retailPrice float ,
                 unit varchar(10),
                 specification varchar(20),
                 lowestRetailPrice float,
                 productNature varchar(15),
                 warrantyPeriod int,
                 distributionMethod varchar(10),
                 estimatedDaysOfUse int,
                 grossProfitMargin float )CHARSET=utf8 COLLATE=utf8_bin;
                 """

# 在 execute里面执行SQL语句
# cursor.execute("CREATE DATABASE product_info_db;")

cursor.execute("USE product_info_db;")
try:
    cursor.execute("DROP TABLE product_info;")
    cursor.execute(sql_createTb)
except Exception as e:
    # 发生错误时回滚
    db.rollback()
    print(str(e))
    exit(-1)
else:
    print("The original table \"product_info\" has been droped！")
    print("New table \"product_info\" has been created!")

colName = table_one.row_values(0)

for i in range(1, nrows):
    # ('{}', '{}', {}, '{}', {}, {}, '{}', '{}', {}, '{}', '{}', '{}', '{}', {})
    lineNum = table_one.cell_value(i, 0)
    barcode = table_one.cell_value(i, 1)
    sellType = table_one.cell_value(i, 2)
    wholesalePrice = table_one.cell_value(1, 3)
    productName = table_one.cell_value(i, 4)
    referencePurchasePrice = table_one.cell_value(i, 5)
    retailPrice = table_one.cell_value(i, 6)
    unit = table_one.cell_value(i, 7)
    specification = table_one.cell_value(i, 8)
    try:
        lowestRetailPrice = float(table_one.cell_value(i, 9))
    except:
        lowestRetailPrice = "null"
    productNature = table_one.cell_value(i, 10)
    try:
        warrantyPeriod = table_one.cell_value(i, 11)
    except:
        warrantyPeriod = "null"
    distributionMethod = table_one.cell_value(i, 12)
    try:
        estimatedDaysOfUse = int(table_one.cell_value(i, 13))
    except:
        estimatedDaysOfUse = "null"
    grossProfitMargin = table_one.cell_value(i, 14)

    # print(barcode, sellType, wholesalePrice, productName, purchasePrice)
    # print(int(str(estimatedDaysOfUse)))

    # ('%s', '%s', % f, '%s', % f, % f, '%s', '%s', % f, '%s', % s, '%s', % s, % f)

    # 将数据存入数据库
    # sql = "insert into product_info(barcode, sellType, wholesalePrice, productName, referencePurchasePrice,retailPrice,unit,specification,lowestRetailPrice,productNature,warrantyPeriod,distributionMethod, estimatedDaysOfUse,grossProfitMargin) values ('%s','%s', %f, '%s', %f, %f,'%s','%s','%d','%s','%s','%s','%s',%f);" % (
    #     barcode, sellType, wholesalePrice, productName, referencePurchasePrice, retailPrice, unit, specification,
    #     lowestRetailPrice, productNature, warrantyPeriod, distributionMethod, estimatedDaysOfUse, grossProfitMargin)
    sql = "insert into product_info(barcode, sellType, wholesalePrice, productName, referencePurchasePrice,retailPrice,unit,specification,lowestRetailPrice,productNature,warrantyPeriod,distributionMethod, estimatedDaysOfUse,grossProfitMargin) values ('{}','{}', {}, '{}', {}, {},'{}','{}',{},'{}',{},'{}',{},{});".format(
        barcode, sellType,
        wholesalePrice, productName,
        referencePurchasePrice,
        retailPrice, unit, specification,
        lowestRetailPrice, productNature,
        warrantyPeriod,
        distributionMethod,
        estimatedDaysOfUse,
        grossProfitMargin)
    print(sql)
    try:
        cursor.execute(sql)
    except Exception as e:
        # 发生错误时回滚
        db.rollback()
        print(str(e))
        exit(-1)
    else:
        db.commit()  # 事务提交
        print('Successful insert row ' + str(i) + '! ')

print("============Finished============")
db.close()
cursor.close()
