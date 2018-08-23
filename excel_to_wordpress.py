# -*- coding: UTF-8 -*-
#!/usr/bin/python
import sys;
if sys.getdefaultencoding() != 'utf-8':
    reload(sys);
    sys.setdefaultencoding('utf-8');

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
    techMaturity="";
    contactName="";
    contactTel="";
    contactEmail="";
    completeDate="";
    applyStatus="";
    cowork_type=""; #合作方式
    apply_score=""; #成果应用行业
    result_type=""; #成果形式
    award_type=""; #获奖类别 10
    award_level="";  #获奖级别 11
    prospect_promotion="";  #推广前景   12
    patent_name="";  #专利名称  13
    postal_address="";  #通讯地址 18
    district="";    #所在地区  19


    for cell in data:
        if cell.has_key(1) :
            title = cell.get(1);
        if( cell.has_key(2)) :
            enterName = cell.get(2);
        if (cell.has_key(3)):
            techField = cell.get(3);
        if cell.has_key(4) :
            content = cell.get(4);
        if (cell.has_key(5)):
            techMaturity = cell.get(5);
        if (cell.has_key(6)):
            applyStatus = cell.get(6);
        if cell.has_key(7) :
            cowork_type = cell.get(7);
        if cell.has_key(8) :
            apply_score = cell.get(8);
        if cell.has_key(9) :
            result_type = cell.get(9);
        if cell.has_key(10):
            award_type = cell.get(10);  # 获奖类别 10
        if cell.has_key(11):
            award_level = cell.get(11);  # 获奖级别 11
        if cell.has_key(12):
            prospect_promotion = cell.get(12);  # 推广前景   12
        if cell.has_key(13):
            patent_name = cell.get(13);  # 专利名称  13
        if (cell.has_key(14)):
            contactName = cell.get(14);
        if (cell.has_key(15)):
            contactTel = cell.get(15);
        if (cell.has_key(16)):
            contactEmail = cell.get(16);
        if cell.has_key( 17) :
            completeDate = cell.get(17);
        if cell.has_key(18):
            postal_address = cell.get(18);  # 通讯地址 18
        if cell.has_key(19):
            district = cell.get(19);  # 所在地区  19
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
    insertOtherDataIntoDB(postId,enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate,
                          cowork_type,district);

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
    print("read excel into wordpress begin.");
    readExcel();
    print("read excel into wordpress end.")