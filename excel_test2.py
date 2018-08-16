#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd;
import MySQLdb;

def excelTest():
    print("excelTest begin,");
    creatTableTest();
    
    workbook = xlrd.open_workbook("""/home/jankin/FTP/example/1.xlsx""");
    #print(workbook.sheet_names());
    for sheetTemp in workbook.sheets() :
        rowNum = sheetTemp.nrows;
        if rowNum != 0 :
            print(sheetTemp.name.encode(encoding='utf-8'));
            printSheet(sheetTemp);

def printSheet(sheetSrc):
    num=0;
    for row in sheetSrc.get_rows() :
        rowList =[];
        b = 0;
        order = 0;
        for cell in row :
            if cell.ctype == xlrd.XL_CELL_TEXT :
                dir = {order:cell.value };
                rowList.append( dir );
                b = 1;
            if cell.ctype == xlrd.XL_CELL_NUMBER :
                dir = {order: str(int(cell.value)) };
                rowList.append(dir);
                b = 1;
            order+=1;
        if  b :
            if len(rowList) >=5 :
                insertDb(rowList);
                num+=1;
                if  num>2 :
                    break;

def insertDb(rowObj):
    print("insert data into DB.");
    insetRow(rowObj);

def creatTableTest():
    print("creat Table Test");
    # 打开数据库连接
    db = MySQLdb.connect("localhost", "root", "peakpeak91", "TESTDB", charset='utf8' );
    # 使用cursor()方法获取操作游标
    cursor = db.cursor();
    # 如果数据表已经存在使用 execute() 方法删除表。
    cursor.execute("DROP TABLE IF EXISTS PROJECT ");

    # 创建数据表SQL语句
    sql = """CREATE TABLE PROJECT (
         ORDER_NAME CHAR(20) NOT NULL,
         PROJECT_NAME  CHAR(50),
         ORG_NAME  CHAR(50),
         SCORE CHAR(50),
         NOTES_TEXT VARCHAR(5000) ) DEFAULT CHARSET=utf8""".encode(encoding='utf-8');
    cursor.execute(sql);
    print("create successful.")
    db.close();

def insetRow(row):
    print("insetRow begin.");
    print(row);

    # 打开数据库连接
    db = MySQLdb.connect("localhost", "root", "peakpeak91", "TESTDB", charset='utf8');
    cursor = db.cursor();
    cursor.execute("SET NAMES utf8");

    sql = """INSERT INTO PROJECT(ORDER_NAME,
             PROJECT_NAME, ORG_NAME, SCORE, NOTES_TEXT) """.encode(encoding='utf-8');
    sql += "VALUES (".encode(encoding='utf-8');
    i =0;
    for cell in row:
        #print(cell);
        if cell.has_key(i):
            value = cell.get(i);
            sql += "\'".encode(encoding='utf-8');
            sql += value.encode(encoding='utf-8');
            sql += "\'".encode(encoding='utf-8');
            if i < 4 :
                sql += ",".encode(encoding='utf-8');
        i+=1;
    sql += ")";
    #print("sql="+sql);

    try:
        cursor.execute(sql);
        db.commit();
    except BaseException as tmp:
        # Rollback in case there is any error
        print("occur except."+str(tmp));
        db.rollback();
    db.close();
    print("insetRow end..");

if __name__ =="__main__" :
    print("excel test begin.");
    excelTest();
    print("excel test end.")