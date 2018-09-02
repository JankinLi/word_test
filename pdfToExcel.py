# -*- coding: UTF-8 -*-
#!/usr/bin/python

from pdfminer.pdfparser import PDFParser;
from pdfminer.pdfparser import PDFDocument;
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter;
from pdfminer.converter import PDFPageAggregator;
from pdfminer.layout import LTTextBoxHorizontal,LAParams;

import xlrd;
import xlwt;

def readPdf():
    print("read pdf from file.");
    pathName = """E:\lichuan\势坤科技\examples\科技成果信息表.pdf""";
    fp = open(pathName, 'rb');
    list = parsePdf(fp);
    fp.close();
    return list;

def parsePdf(fp):
    print("parsePdf begin.");
    praser = PDFParser(fp);
    doc = PDFDocument();
    praser.set_document(doc);
    doc.set_parser(praser);

    doc.initialize();

    if not doc.is_extractable:
        print("document is not extractable");
        return;

    rsrcmgr = PDFResourceManager();
    laparams = LAParams();
    device = PDFPageAggregator(rsrcmgr, laparams=laparams);

    interpreter = PDFPageInterpreter(rsrcmgr, device);
    i=0;
    list=[];
    for page in doc.get_pages():
        interpreter.process_page(page);
        layout = device.get_result();
        results="";

        for x in layout:
            if (isinstance(x, LTTextBoxHorizontal)):
                results += x.get_text();
        i+=1;
        if i>10 :
            list.append(results);
            #if i > 41:
            #   break;

    #for table in list:
    #    print(table);
    #    print("======================\n");
    return list;

def writeExcel(list):
    print("writeExcel begin.");
    #book = xlwt.Workbook(encoding='utf-8', style_compression=0);
    book = xlwt.Workbook(encoding='utf-8');
    sheet = book.add_sheet('test', cell_overwrite_ok=True);

    sheet.write(0,2 ,"成果名称");
    sheet.write(0,3 ,"所属单位");
    sheet.write(0,4 ,"联 系 人");
    sheet.write(0,5 ,"联系电话");
    sheet.write(0,6 ,"成果简介");

    i=1;
    for str in list :
        writeLine(i, str, sheet);
        i+=1;

    book.save(r"""E:\lichuan\势坤科技\examples\pdf_1.xls""");
    print("writeExcel end.");

def writeLine(i, line, sheet):
    sheet.write(i, 1, "page " + str(i));
    j=2;
    print(line);
    array = line.split('\n',10);
    print(len(array));

    order = 1;
    type = 0;
    for val in array:
        if (val == "成果名称") and (order == 1):
            type = 1;
        if (val == "成果信息表") and (order == 1):
            type = 2;

        if (type == 2) and (val == "成果名称") and (order == 2):
            type = 3;

        if (type == 3) and (val == "联 系 人") and (order == 4):
            type = 4;

        if (type == 1 and ( order == 4 or order == 5 or order == 7 or order == 9 or order == 11 ) ):
            sheet.write(i, j, val);
            j += 1;
        if (type == 2 and ( order == 2 or order == 3 or order == 4 or order == 6 or order == 11 ) ):
            sheet.write(i, j, val);
            j += 1;
        if (type == 3 and ( order == 4 or order == 5 or order == 7 or order == 9 or order == 11 ) ):
            sheet.write(i, j, val);
            j += 1;
        if (type == 4 and ( order == 5 or order == 6 or order == 7 or order == 9 or order == 11 ) ):
            sheet.write(i, j, val);
            j += 1;
        order+=1;

def readTest():
    workbook = xlrd.open_workbook(r"""E:\lichuan\势坤科技\examples\pdf_1.xls""");
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
            order += 1;
        print(rowList);

if __name__ == "__main__":
    print("begin");
    list = readPdf();
    #print(list);
    writeExcel(list);
    #readTest();
    print("end.");