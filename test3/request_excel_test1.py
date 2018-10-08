# -*- coding: UTF-8 -*-
#!/usr/bin/python
import sys;
if sys.getdefaultencoding() != 'utf-8':
    reload(sys);
    sys.setdefaultencoding('utf-8');

import os;
import xlrd;
import MySQLdb;
from wordpress_xmlrpc import Client, WordPressPost;
from wordpress_xmlrpc.methods.posts import NewPost;

def readExcel():
    print("readExcel begin,");
    workbook = xlrd.open_workbook(r"""/home/jankin/examples/2018年沈阳市科技型企业创新需求库.xlsx""");
    #workbook = xlrd.open_workbook(r"""E:\lichuan\势坤科技\examples\2018年沈阳市科技型企业创新需求库.xlsx""");
    for sheetTemp in workbook.sheets() :
        rowNum = sheetTemp.nrows;
        if rowNum != 0 :
            if( sheetTemp.name == r"""需求表（需）"""):
                readRequest(sheetTemp);

def readRequest(sheetSrc):
    print("read request begin.");
    print("row:"+str(sheetSrc.nrows));
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
                filterData(rowList);
                num += 1;

    print("num="+str(num));

def filterData(data):
    for item in data:
        if 0 in item.keys() :
            val = item[0];
            if val.isdigit() :
                createPost(data);

def createPost(list):
    #print( list );
    #print( "create post begin.");

    title = "";  #标题
    enterName = ""; #企业名称
    field1 = ""; #所属一级领域
    field2 = ""; #所属二级领域
    techField = ""; #技术领域
    reason_1 = ""; # 是否高企
    reason_2 = "";  # 2017双培育入库
    reason_3 = "";  # 2018双培育申报
    reason_4 = "";  # 是否承担项目
    addReason = "";  # 入选理由
    contactName = "";#联系人
    contactTel = "";
    deparment = ""; #责任部门
    enterType= ""; #企业性质
    bigMoney = "";  # 千亿产业链
    income_2017 = "";  # 2017年营业收入
    outcomt_2017 = "";  # 2017年研发投入
    create_request = "";  # 创新需求
    other_reqest = "";  # 其他需求
    dis_address = "";  # 所属区县
    district = "沈阳"; # 所在地区

    for item in list:
        if 1 in item.keys() :
            enterName = item[1];
            title = item[1] + "的需求";
        if 2 in item.keys() :
            dis_address = item[2];
        if 3 in item.keys() :
            field1 = item[3];
        if 4 in item.keys() :
            field2 = item[4];
        if 5 in item.keys() :
            reason_1 = item[5];
        if 6 in item.keys() :
            reason_2 = item[6];
        if 7 in item.keys() :
            reason_3 = item[7];
        if 8 in item.keys() :
            reason_4 = item[8];
        if 9 in item.keys() :
            contactName = item[9];
        if 10 in item.keys() :
            contactTel = item[10];
        if 11 in item.keys() :
            deparment = item[11];
        if 12 in item.keys():
            enterType = item[12];
        if 13 in item.keys():
            bigMoney = item[13];
        if 14 in item.keys():
            income_2017 = item[14];  # 2017年营业收入
        if 15 in item.keys():
            outcomt_2017 = item[15];  # 2017年研发投入
        if 16 in item.keys():
            create_request = item[16]; # 创新需求
        if 17 in item.keys():
            other_reqest = item[17];  # 其他需求
    techField = genTechField(field1, field2);
    addReason = genReason(reason_1, reason_2, reason_3, reason_4);
    #print( title + "," + techField + "," + addReason );
    postNewPostByXmlRpc(title,techField,addReason,enterName ,contactName, contactTel,
        deparment, enterType, bigMoney, income_2017, outcomt_2017, create_request,
        other_reqest ,  dis_address ,  district );

def postNewPostByXmlRpc(title, techField,addReason,enterName ,contactName, contactTel,
        deparment, enterType, bigMoney, income_2017, outcomt_2017, create_request,
        other_reqest , dis_address ,  district ):
    print("postNewPostByXmlRpc");
    wp = Client('http://39.106.104.45/wordpress/xmlrpc.php', 'shikun', 'ShiKun001')
    post = WordPressPost();
    post.title = title;
    post.content = gen_content(enterType,addReason,bigMoney, income_2017, outcomt_2017, create_request, other_reqest , dis_address);
    post.terms_names = {
        'category': ['企业需求']
    };
    post.id = wp.call(NewPost(post));
    print("post.id = " + str(post.id));
    postId = post.id;
    insertOtherDataIntoDB(postId,enterName,techField,deparment,contactName,contactTel,district);
    insertProvDataIntoDB(postId);
    savePostId(postId);

def gen_content(enterType,addReason,bigMoney, income_2017, outcomt_2017, create_request, other_reqest , dis_address):
    str="""<h3>企业需求</h3>""".encode(encoding='utf-8');
    str+="""<h4>创新需求</h4>""".encode(encoding='utf-8');
    str +="""<p>""".encode(encoding='utf-8');
    str += create_request;
    str += """</p>""".encode(encoding='utf-8');
    str += """<h4> 人才 / 资金等需求 </h4>""".encode(encoding='utf-8');
    str += """<p>""".encode(encoding='utf-8');
    str += other_reqest;
    str += """</p>""".encode(encoding='utf-8');
    str += """<h4> 企业信息 </h4>""".encode(encoding='utf-8');
    str += """<p> 2017 年营业收入(万元)：""".encode(encoding='utf-8');
    str += income_2017;
    str += """<br>2017 年研发投入(万元)：""".encode(encoding='utf-8');
    str += outcomt_2017;
    str += """</p>""".encode(encoding='utf-8');
    str += """<p> 企业性质：""".encode(encoding='utf-8');
    str += enterType;
    str += """<br>入选理由：""".encode(encoding='utf-8');
    str += addReason;
    str += """<br>千亿产业链：""".encode(encoding='utf-8');
    str += bigMoney;
    str += """</p>""".encode(encoding='utf-8');
    str += """<p> 所属区县：""".encode(encoding='utf-8');
    str += dis_address;
    str += """</p>""".encode(encoding='utf-8');
    return str;

def insertOtherDataIntoDB(postId,enterName,techField,deparment,contactName,contactTel,district):
    print("insertOtherDataIntoDB, postId= " + str(postId));
    if len(enterName) > 0:
        insertMeta(postId, "enter-name", enterName);
    if len(techField)>0 :
        insertMeta(postId, "tech-field", techField);
    contectVal = contactName;
    if len(deparment)> 0 :
        contectVal = deparment + " : " + contactName;
    if len(contectVal) > 0:
        insertMeta(postId, "contact-name", contectVal);
    if len(contactTel) > 0:
        insertMeta(postId, "contact-tel", contactTel);
        print("contact-tel=" + contactTel);
    if len(district)>0 :
        insertMeta(postId, "district_name", district);

def insertMeta(postId, key, value):
    #print("insertMeta begin.")
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
    #print("insertMeta end..");

def insertProvDataIntoDB(postId,):
    print("insertProvDataIntoDB , postId="+str(postId));
    code = 36;
    parentCode = 35;

    db = MySQLdb.connect("localhost", "root", "magic123", "bitnami_wordpress", charset='utf8',
                         unix_socket='/opt/wordpress-4.9.8-0/mysql/tmp/mysql.sock');
    cursor = db.cursor();
    cursor.execute("SET NAMES utf8");

    ret = insertDataIntoTermRelationShip(db,cursor, postId, str(code));
    if ret == 0 :
        insertDataIntoTermRelationShip(db, cursor, postId, str(parentCode));
    db.close();

def insertDataIntoTermRelationShip(db, cursor, postId, code):
    sql = """INSERT INTO wp_term_relationships(object_id,
                            term_taxonomy_id, term_order) """.encode(encoding='utf-8');
    sql += "VALUES (".encode(encoding='utf-8');
    sql += postId;
    sql += ",";
    sql += code;
    sql += ",0";
    sql += ")";
    # print("sql="+sql);

    try:
        cursor.execute(sql);
        db.commit();
    except BaseException as tmp:
        # Rollback in case there is any error
        print("occur except." + str(tmp));
        db.rollback();
        return 1;
    return 0;

def savePostId(postId):
    f = open("/opt/tmp/post_id_request.txt", "a");
    f.write(str(postId));
    f.write('\n');
    f.close();

def genTechField(f1, f2):
    temp = str(f1) + " / " + str(f2);
    return temp;

def genReason(reason_1,reason_2,reason_3,reason_4):
    temp = "是否是高新技术企业(";
    if len(reason_1) >0 :
        temp += reason_1;
        temp += "),";
    else:
        temp += "否),";
    temp += "是2017双培育入库企业(";
    if len(reason_2) >0 :
        temp += reason_2;
        temp += "),";
    else:
        temp += "否),";
    temp += "是2018双培育申报企业(";
    if len(reason_3) > 0:
        temp += reason_3;
        temp += "),";
    else:
        temp += "否),";
    temp += "是否是承担项目的企业(";
    if len(reason_4) > 0:
        temp += reason_4;
        temp += "),";
    else:
        temp += "否),";
    return temp;

if __name__ == '__main__' :
    print("request begin");

    try:
        os.remove("""/opt/tmp/post_id_request.txt""");
    except OSError as tmp:
        print(tmp);

    readExcel();
    print("request end.");
