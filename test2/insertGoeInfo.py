# -*- coding: UTF-8 -*-
#!/usr/bin/python

import MySQLdb;

def insertMeta(postId, goe, value):
    print("insertMeta begin.")
    db = MySQLdb.connect("localhost", "root", "magic123", "bitnami_wordpress", charset='utf8', unix_socket='/opt/wordpress-4.9.8-0/mysql/tmp/mysql.sock');
    cursor = db.cursor();
    cursor.execute("SET NAMES utf8");

    sql = """INSERT INTO wp_term_relationships(object_id,
                 term_taxonomy_id, term_order) """.encode(encoding='utf-8');
    sql += "VALUES (".encode(encoding='utf-8');
    sql += postId;
    sql += ",";
    sql += str(goe);
    sql += ",";
    sql += str(value);
    sql += ")";
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

def readAllId(list):
    f = open("/opt/tmp/postId.txt", "r");
    for line in f:
        val = line.strip('\n');
        list.append(val);
    f.close();

if __name__ =="__main__" :
    print("insert goe begin.");
    list=[];
    readAllId(list);
    print(list);
    for item in list:
        if item != '1077':
            insertMeta(item,27,0);
    print("insert goe end.");