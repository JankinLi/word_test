# -*- coding: UTF-8 -*-
#!/usr/bin/python

import xlwt;
import xml.sax;

class XmlHanlder ( xml.sax.ContentHandler):
    def __init__(self):
        self.CurrentData = "";
        self.tableName = "";
        self.title ="";
        self.orgName ="";
        self.contactName="";
        self.phoneVal="";
        self.content="";
        self.pageStart= False;

    # 元素开始事件处理
    def startElement(self, tag, attributes):
        self.CurrentData = tag;
        if tag == "page" :
            self.pageStart = True;
            print("page");

    # 元素结束事件处理
    def endElement(self, tag):
        if tag == "page":
            self.pageStart = False;

    # 内容事件处理
    def characters(self, content):
        print(content);

def readXml(path):
    print("readXml begin.path="+path);
    # 创建一个 XMLReader
    parser = xml.sax.make_parser();
    # turn off namepsaces
    parser.setFeature(xml.sax.handler.feature_namespaces, 0);
    Handler = XmlHanlder();
    parser.setContentHandler(Handler);
    parser.parse(path);

if __name__ == "__main__":
    print("begin");
    readXml("""/home/jankin/examples/xml/1/1.xml""");
    print("end.");