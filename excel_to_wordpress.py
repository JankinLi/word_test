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
                #print(rowList);
                num += 1;
                if num > 1:
                    postNewPost(num, rowList);
                    if( num > 3) :
                        break;

def postNewPost(order, data):
    print("postNewPost begin. order="+ str(order));
    if (len(data) < 5):
        return;
    for cell in data:
        if cell.has_key(1) :
            title = cell.get(1);
        if cell.has_key(4) :
            content = cell.get(4);
    postNewPostByXmlRpc(title, content);

def postNewPostByXmlRpc(title, content):
    print("postNewPostByXmlRpc");
    wp = Client('http://39.106.104.45/wordpress/xmlrpc.php', 'shikun', 'ShiKun001')
    post = WordPressPost();
    post.title = title;
    post.content = content;
    post.terms_names = {
        'post_tag': ['test', 'lichuan'],
        'category': ['Introductions', 'Tests']
    };
    post.id = wp.call(NewPost(post));
    print("post.id = " + str(post.id));

if __name__ =="__main__" :
    print("read excel into wordpress begin.");
    readExcel();
    print("read excel into wordpress end.")