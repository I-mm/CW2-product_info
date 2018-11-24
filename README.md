# Coursework 2: Product Info

- [x] Input the .xlsx file into mysql database.
- [x] Deal with the unreasonable errors in the table
- [x] Deal with the NULL data.
- [x] Analyse and visualize the data with SQL statement, Pandas and Ｍatplotlib. 
- [x] Count the profit margin of the product and give the sales strategy.

<br>

### Dependencies 

```
conda env create -f environment.yml
```

<br>

### Data Input (with preliminary screening of data)

See [excelToDB.py](https://github.com/I-mm/CW2-product_info/blob/master/excelToDB.py). 

- Connect remote mysql database:

```python
db = pymysql.connect(host="39.105.165.114", user="root", password="zym2112!", use_unicode=True, charset="utf8")
```
- Create table product_info:

```sql
CREATE TABLE product_info (
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
```

- Insert statement

```python
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
```

<br>

### Data Analysis
See [data analysis.ipynb](https://github.com/I-mm/CW2-product_info/blob/master/data_analysis.ipynb). 

<br>

### Contributor

[@赵屹铭](https://github.com/I-mm)

