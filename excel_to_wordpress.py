#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd;
import MySQLdb;
from wordpress_xmlrpc import Client, WordPressPost;
from wordpress_xmlrpc.methods.posts import NewPost;

def readExcel():
    print("readExcel begin,");
    workbook = xlrd.open_workbook("""/home/jankin/examples/北京理工大学科技成果.xlsx""");
    for sheetTemp in workbook.sheets() :
        rowNum = sheetTemp.nrows;
        if rowNum != 0 :
            parseData(sheetTemp);

def parseData(sheetSrc):
    num = 0;
    for row in sheetSrc.get_rows():
        rowList = [];
        b = 0;
        order = 0;
        for cell in row:
            if cell.ctype == xlrd.XL_CELL_TEXT:
                dir = {order: cell.value};
                rowList.append(dir);
                b = 1;
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                dir = {order: str(int(cell.value))};
                rowList.append(dir);
                b = 1;
            order += 1;
        if b:
            if len(rowList) >= 5:
                print(rowList);
                num += 1;
                if num > 1:
                    postNewPost(rowList);
                    break;

def postNewPost(data):
    print("postNewPost begin.");
    

if __name__ =="__main__" :
    print("read excel into wordpress begin.");
    readExcel();
    print("read excel into wordpress end.")