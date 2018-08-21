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
                    #if( num > 3) :
                    break;

def postNewPost(order, data):
    print("postNewPost begin. order="+ str(order));
    if (len(data) < 5):
        return;
    title="";
    content="";
    enterName="";
    techField="";
    contactName="";
    contactTel="";
    contactEmail="";
    completeDate="";
    for cell in data:
        if cell.has_key(1) :
            title = cell.get(1);
        if cell.has_key(4) :
            content = cell.get(4);
        if( cell.has_key(2)) :
            enterName = cell.get(2);
        if (cell.has_key(3)):
            techField = cell.get(3);
        if (cell.has_key(5)):
            techMaturity = cell.get(5);
        if (cell.has_key(14)):
            contactName = cell.get(14);
        if (cell.has_key(15)):
            contactTel = cell.get(15);
        if (cell.has_key(16)):
            contactEmail = cell.get(16);
        if cell.has_key( 17) :
            completeDate = cell.get(17);
    postNewPostByXmlRpc(title, content,enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate);

def postNewPostByXmlRpc(title, content, enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate):
    print("postNewPostByXmlRpc");
    wp = Client('http://39.106.104.45/wordpress/xmlrpc.php', 'shikun', 'ShiKun001')
    post = WordPressPost();
    post.title = title;
    post.content = content;
    post.terms_names = {
        'post_tag': ['北京理工大学'],
        'category': ['成果展示']
    };
    #post.custom_fields = {
    #    'enter-name':enterName,
    #    'tech-field': techField,
    #    'tech-maturity': techMaturity,
    #    'contact-name':contactName,
    #    'contact-tel':contactTel,
    #    'contact-email':contactEmail
    #};
    post.id = wp.call(NewPost(post));
    print("post.id = " + str(post.id));
    postId = post.id;
    insertOtherDataIntoDB(postId,enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate);

def insertOtherDataIntoDB(postId,enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate) :
    print("insertOtherDataIntoDB  postId= " + str(postId));
    insertMeta(postId, "enter-name", enterName);
    insertMeta(postId, "tech-field", techField);
    insertMeta(postId, "tech-maturity", techMaturity);
    insertMeta(postId, "contact-name", contactName);
    insertMeta(postId, "contact-tel", contactTel);
    print("contact-tel=" + contactTel);
    insertMeta(postId, "contact-email", contactEmail);
    if len(completeDate)> 0 :
        print("tech-date="+completeDate);
        inserMeta(postId, "tech-date", completeDate);

def insertMeta(postId, key, value):
    print("insertMeta begin.")
    db = MySQLdb.connect("localhost", "root", "magic123", "bitnami_wordpress", charset='utf8', unix_socket='/opt/wordpress-4.9.8-0/mysql/tmp/mysql.sock');
    cursor = db.cursor();
    cursor.execute("SET NAMES utf8");

    sql = """INSERT INTO wp_postmeta(post_id,
                 meta_key, meta_value) """.encode(encoding='utf-8');
    sql += "VALUES (".encode(encoding='utf-8');
    sql += postId;
    sql += ",'";
    sql += key;
    sql += "','";
    sql += value;
    sql += "')";
    # print("sql="+sql);

    try:
        cursor.execute(sql);
        db.commit();
    except BaseException as tmp:
        # Rollback in case there is any error
        print("occur except." + str(tmp));
        db.rollback();
    db.close();
    print("insertMeta end..");

if __name__ =="__main__" :
    print("read excel into wordpress begin.");
    readExcel();
    print("read excel into wordpress end.")