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
    workbook = xlrd.open_workbook("""/opt/temp/pdf_2.xls""");
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
            if len(rowList) > 5:
                #print(rowList);
                num += 1;
                if num >= 1:
                    postNewPost(num, rowList);
                    #if( num >= 2) :
                    #    break;

def postNewPost(order, data):
    print("postNewPost begin. order="+ str(order) + ",data="+ str(data) );
    if (len(data) < 5):
        return;

    title = "";
    content = "";
    enterName = "";
    techField = "";
    techMaturity = "";
    contactName = "";
    contactTel = "";
    contactEmail = "";
    completeDate = "";
    applyStatus = "";
    cowork_type = "";  # 合作方式
    apply_score = "";  # 成果应用行业
    result_type = "";  # 成果形式
    award_type = "";  # 获奖类别 10
    award_level = "";  # 获奖级别 11
    prospect_promotion = "";  # 推广前景   12
    patent_name = "";  # 专利名称  13
    postal_address = "";  # 通讯地址 18
    district = "";  # 所在地区  19

    for cell in data:
        if cell.has_key(2) :
            title = cell.get(2);
        if( cell.has_key(3)) :
            enterName = cell.get(3);
        if (cell.has_key(4)):
            contactName = cell.get(4);
        if (cell.has_key(5)):
            contactTel = cell.get(5);
        if cell.has_key(6) :
            content = cell.get(6);

    postNewPostByXmlRpc(title, content,enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate,
                        applyStatus,cowork_type,apply_score,result_type,award_type,award_level,prospect_promotion,
                        patent_name,postal_address,district);

def postNewPostByXmlRpc(title, content, enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate,
                        applyStatus, cowork_type, apply_score, result_type, award_type, award_level, prospect_promotion,
                        patent_name, postal_address, district):
    print("postNewPostByXmlRpc");
    wp = Client('http://39.106.104.45/wordpress/xmlrpc.php', 'shikun', 'ShiKun001')
    post = WordPressPost();
    post.title = title;
    post.content = gen_content(content, enterName,applyStatus,apply_score, result_type, award_type, award_level,
                               prospect_promotion,patent_name,postal_address);
    post.terms_names = {
        'category': ['成果展示']
    };

    post.id = wp.call(NewPost(post));
    print("post.id = " + str(post.id));
    postId = post.id;
    insertOtherDataIntoDB(postId,enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate,
                          cowork_type,district);
    savePostId(postId);

def savePostId(postId):
    f = open("/opt/tmp/post_id_pdf.txt", "a");
    f.write(str(postId));
    f.write('\n');
    f.close();

def gen_content(content, enterName,applyStatus,apply_score, result_type, award_type, award_level, prospect_promotion,
                patent_name,postal_address):
    str="""<h3>成果介绍</h3>""".encode(encoding='utf-8');
    str+="""<p>""".encode(encoding='utf-8');
    str+=content;
    str+="""</p>""".encode(encoding='utf-8');
    str+="""<hr />""".encode(encoding='utf-8');
    str+="""<table border="1" cellspacing="0" cellpadding="5">
<tbody>
<tr>""".encode(encoding='utf-8');
    str += """<td><b>完成单位:</b>""".encode(encoding='utf-8');
    str+=enterName;
    str+="""</td>
<td><b>应用情况:</b>""".encode(encoding='utf-8');
    str+=applyStatus;
    str+="""</td>
</tr>
<tr>""".encode(encoding='utf-8');
    str+="""<td><b>获奖类别:</b>""".encode(encoding='utf-8');
    str+=award_type;
    str+="""</td>
<td><b>获奖级别:</b>""".encode(encoding='utf-8');
    str+=award_level;
    str+="""</td>
</tr>
<tr>
<td colspan="2"><b>成果应用行业:</b>""".encode(encoding='utf-8');
    str+=apply_score;
    str+="""<br>
</td>
</tr>
<tr>
<td colspan="2"><b>成果形式:</b><br>""".encode(encoding='utf-8');
    str+=result_type;
    str+="""</td>
</tr>
<tr>
<td colspan="2"><b>推广前景:</b><br>""".encode(encoding='utf-8');
    str+=prospect_promotion;
    str+="""</td>
</tr>
<tr>
<td colspan="2"><b>专利名称:</b><br>""".encode(encoding='utf-8');
    str += patent_name;
    str += """</td>
    </tr>
    <tr>
    <td colspan="2"><b>通讯地址:</b><br>""".encode(encoding='utf-8');
    str += postal_address;
    str += """</td>
    </tr>
</tbody>
</table>
    """.encode(encoding='utf-8');
    return str;

def insertOtherDataIntoDB(postId,enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate,
                          cowork_type, district) :
    print("insertOtherDataIntoDB  postId= " + str(postId));
    if len(enterName) > 0:
        insertMeta(postId, "enter-name", enterName);
    if len(techField)>0 :
        insertMeta(postId, "tech-field", techField);
    if len( techMaturity) > 0 :
        insertMeta(postId, "tech-maturity", techMaturity);
    if len(contactName) > 0:
        insertMeta(postId, "contact-name", contactName);
    if len(contactTel) > 0:
        insertMeta(postId, "contact-tel", contactTel);
        print("contact-tel=" + contactTel);
    if len(contactEmail)>0:
        insertMeta(postId, "contact-email", contactEmail);
    if len(completeDate)> 0 :
        print("tech-date="+completeDate);
        insertMeta(postId, "tech-date", completeDate);
    if len(cowork_type)>0:
        insertMeta(postId, "cowork-type", cowork_type);
    if len(district)>0 :
        insertMeta(postId, "district_name", district);

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
    print("read pdf_2.xls excel into wordpress begin.");
    try:
        os.remove("""/opt/tmp/post_id_pdf.txt""");
    except OSError as tmp:
        print(tmp);

    readExcel();

    print("read pdf_2.xls end.")
