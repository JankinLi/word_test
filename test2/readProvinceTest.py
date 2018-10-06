# -*- coding: UTF-8 -*-
#!/usr/bin/python

import xlrd;

def readProvince(path):
    list=[];
    workbook = xlrd.open_workbook(path);
    for sheetTemp in workbook.sheets() :
        rowNum = sheetTemp.nrows;
        if rowNum != 0:
            print(rowNum);
            list = readSheet(sheetTemp);
            break;
    return list;

def readSheet(sheetSrc):
    num = 0;
    result = [];
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
        if b == 1:
            print(rowList);
            result.append(rowList);
    return result;

def parseResult(list):
    ret = [];
    for row in list:
        dir ={};
        for item in row:
            if 1 in item :
                #print (item.get(1));
                dir[1]=item.get(1);
            if 3 in item:
                #print( item.get(3));
                dir[3] = item.get(3);
            if 4 in item:
                #print( item.get(4));
                dir[4] = item.get(4);
        ret.append(dir);
    return ret;

if __name__ == "__main__":
    print("begin");
    #result = readProvince(r"""/home/jankin/examples/province_pinyin_1006.xlsx""");
    result = readProvince(r"""E:\lichuan\势坤科技\doc\province_pinyin_1006.xlsx""");
    if len(result)>0:
        print (result);
        ret = parseResult(result);
        print (ret);
    print("end");