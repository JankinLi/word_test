#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd;

def excelTest():
    print("excelTest begin,");
    workbook = xlrd.open_workbook("""E:\lichuan\势坤科技\examples\北京理工大学科技成果.xlsx""");
    #print(workbook.sheet_names());
    for sheetTemp in workbook.sheets() :
        rowNum = sheetTemp.nrows;
        if rowNum != 0 :
            print(sheetTemp.name);
            printSheet(sheetTemp);

def printSheet(sheetSrc):
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
                dir = {order:str(int(cell.value)) };
                rowList.append(dir);
                b = 1;
            order+=1;
        if  b :
            print(rowList);

if __name__ =="__main__" :
    print("excel test begin.");
    excelTest();
    print("excel test end.")