# -*- coding: UTF-8 -*-
#!/usr/bin/python

import xlwt;
import xml.sax;

class XmlHanlder ( xml.sax.ContentHandler):
    def __initData(self):
        self.tableName = "";
        self.title = "";
        self.orgName = "";
        self.contactName = "";
        self.phoneVal = "";
        self.content = "";
        self.appendContent = "";
        self.images=[];

    def __init__(self):
        self.CurrentData = "";
        self.__initData();
        self.pageStart= False;
        self.pageNumher = 0;
        self.list = [];
        self.hasBegin= False;
        self.flag = 0;

    # 元素开始事件处理
    def startElement(self, tag, attributes):
        self.CurrentData = tag;
        if tag == "page" :
            self.pageStart = True;
            self.pageNumher = attributes["number"];
            print("page " + str(self.pageNumher));
        if self.CurrentData == "image":
            self.images.append( attributes["src"] );

    # 元素结束事件处理
    def endElement(self, tag):
        if int(self.pageNumher) < 11 :
            return;
        if tag == "page":
            self.flag = -1;
            self.pageStart = False;
            if self.hasBegin:
                d = {};
                d[1] = self.title;
                d[2] = self.orgName;
                d[3] = self.contactName;
                d[4] = self.phoneVal;
                d[5] = self.content;
                if len(self.images) >0 :
                    d[6] = self.images;
                self.list.append(d);
                self.hasBegin=False;
            else:
                if len(self.appendContent)>0 :
                    lastItem = self.list[len(self.list)-1];
                    temp= lastItem[5];
                    temp+=self.appendContent;
                    lastItem[5] = temp;
                if len(self.images) >0 :
                    lastItem = self.list[len(self.list) - 1];
                    if 6 in lastItem.keys() :
                        temp = lastItem[6];
                        for img in self.images:
                            temp.append(img);
                        lastItem[6] = temp;
                    else:
                        lastItem[6] = self.images;
            self.__initData();

    # 内容事件处理
    def characters(self, content):
        if int(self.pageNumher) < 11 :
            return;
        if content == '\n':
            return;
        if self.CurrentData == "text" :
            print( content );
            if content == "成果信息表" :
                self.hasBegin = True;
                return;
            if content == "成果名称":
                self.flag = 1;
                return;
            if content == "所属单位":
                self.flag = 2;
                return;
            if content == "联 系 人":
                self.flag = 3;
                return;
            if content == "联系电话":
                self.flag = 4;
                return;

            if content == "成果简介":
                self.flag = 5;
                return;

            if self.hasBegin :
                if self.flag == 1 :
                    self.title = content;
                    return;
                if self.flag == 2:
                    self.orgName = content;
                    return;

                if self.flag == 3 :
                    self.contactName = content;
                    return;

                if self.flag == 4:
                    self.phoneVal = content;
                    return;

                if self.flag == 5:
                    self.content += content;
                    return;
            else:
                if content.isdecimal() :
                    if int(content) == int(self.pageNumher) :
                        return;
                self.appendContent += content;




def readXml(path):
    print("readXml begin.path="+path);
    # 创建一个 XMLReader
    parser = xml.sax.make_parser();
    # turn off namepsaces
    parser.setFeature(xml.sax.handler.feature_namespaces, 0);
    Handler = XmlHanlder();
    parser.setContentHandler(Handler);
    parser.parse(path);
    print("Page Number:"+str(Handler.pageNumher));
    i = 0;
    for item in Handler.list:
        print (item);
        i+=1;
    print("sum:" +str(i));
    return Handler.list;

def writeExcel(list):
    print("writeExcel begin.");
    book = xlwt.Workbook(encoding='utf-8');
    sheet = book.add_sheet('test', cell_overwrite_ok=True);

    sheet.write(0,2 ,"成果名称");
    sheet.write(0,3 ,"所属单位");
    sheet.write(0,4 ,"联 系 人");
    sheet.write(0,5 ,"联系电话");
    sheet.write(0,6 ,"成果简介");
    sheet.write(0, 7, "图片");

    i = 1;
    for item in list:
        writeLine(i, item, sheet);
        i += 1;

    book.save(r"""E:\lichuan\势坤科技\temp\pdf_to_xml_1.xls""");
    print("writeExcel end.");

def writeLine(i, item, sheet):
    sheet.write(i, 1, "page " + str(i));
    j=2;
    for key in item.keys():
        val = item[key];
        if key == 6 :
            path = "";
            for valTemp in val:
                path+=str(valTemp);
                path+="|";
            sheet.write(i,j,path);
        else:
            sheet.write(i, j, val);
        j+=1;

def writeImageTxt(list):
    f = open(r"""E:\lichuan\势坤科技\temp\image.txt""", "w");
    for item in list:
        writeImageItem(item, f);

    f.close();

def writeImageItem(item, f):
    for key in item.keys():
        val = item[key];
        if key == 6 :
            for valTemp in val:
                f.write(valTemp);
                f.write('\n');

if __name__ == "__main__":
    print("begin");
    result=[];
    result = readXml(r"""E:\lichuan\势坤科技\examples\1\1.xml""");
    writeExcel(result);
    writeImageTxt(result);
    print("end.");