# -*- coding: UTF-8 -*-
#!/usr/bin/python

import sys;
if sys.getdefaultencoding() != 'utf-8':
    reload(sys);
    sys.setdefaultencoding('utf-8');

import os;
import xlrd;
import xml.sax;
import MySQLdb;
from wordpress_xmlrpc import Client, WordPressPost;
from wordpress_xmlrpc.methods.posts import NewPost;

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
            #print("page " + str(self.pageNumher));
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
            #print( content );
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
    #print("Page Number:"+str(Handler.pageNumher));
    i = 0;
    for item in Handler.list:
        #print (item);
        i+=1;
    print("sum:" +str(i));
    return Handler.list;

def createAllPost(list,prov):
    print("createAllPost size="+ str(len(list)));
    num = 0;
    for item in list:
        createPostByItem(item,prov);
        num+=1;
        #if num >2 :
        #    break;
    print("num="+str(num));

def createPostByItem(item,prov):
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
    images=[];

    if item.has_key(1):
        title = item.get(1);
    if (item.has_key(2)):
        enterName = item.get(2);
    if (item.has_key(3)):
        contactName = item.get(3);
    if item.has_key(4):
        contactTel= item.get(4);
    if item.has_key(5):
        content = item.get(5);

    if item.has_key(6):
        images = item.get(6);

    if len(enterName) >0 :
        enterTempName = enterName[0:2];
        codeVal = findCode(prov,enterTempName);
        if codeVal != 89 :
            district = enterTempName;

    postNewPostByXmlRpc(title, content, enterName, techField, techMaturity, contactName, contactTel, contactEmail,
                        completeDate,
                        applyStatus, cowork_type, apply_score, result_type, award_type, award_level, prospect_promotion,
                        patent_name, postal_address, district,
                        images,prov);

def postNewPostByXmlRpc(title, content, enterName,techField,techMaturity,contactName,contactTel,contactEmail,completeDate,
                        applyStatus, cowork_type, apply_score, result_type, award_type, award_level, prospect_promotion,
                        patent_name, postal_address, district,
                        images,prov):
    print("postNewPostByXmlRpc");
    wp = Client('http://39.106.104.45/wordpress/xmlrpc.php', 'shikun', 'ShiKun001')
    post = WordPressPost();
    post.title = title;
    post.content = gen_content(content, enterName,applyStatus,apply_score, result_type, award_type, award_level,
                               prospect_promotion,patent_name,postal_address,images);
    post.terms_names = {
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

    insertProvDataIntoDB(postId, prov, enterName);
    savePostId(postId);

def savePostId(postId):
    f = open("/opt/tmp/post_id_from_xml.txt", "a");
    f.write(str(postId));
    f.write('\n');
    f.close();

def gen_content(content, enterName,applyStatus,apply_score, result_type, award_type, award_level, prospect_promotion,
                patent_name,postal_address,images):
    str="""<h3>成果介绍</h3>""".encode(encoding='utf-8');
    str+="""<p>""".encode(encoding='utf-8');
    str+=content;
    str+="""</p>""".encode(encoding='utf-8');
    if len(images)>0:
        for imgPath in images:
            str+="""<img src = "http://39.106.104.45/wordpress/wp-content/uploads/2018/10/""".encode(encoding='utf-8');
            str+=imgPath;
            str+=""""/>""";
            str+="\n";

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
        #print("contact-tel=" + contactTel);
    if len(contactEmail)>0:
        insertMeta(postId, "contact-email", contactEmail);
    if len(completeDate)> 0 :
        #print("tech-date="+completeDate);
        insertMeta(postId, "tech-date", completeDate);
    if len(cowork_type)>0:
        insertMeta(postId, "cowork-type", cowork_type);
    if len(district)>0 :
        insertMeta(postId, "district_name", district);

def insertMeta(postId, key, value):
    #print("insertMeta begin.");
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

def insertProvDataIntoDB(postId, prov, enterName):
    print("insertProvDataIntoDB begin.postId="+str(postId));

    prefix = enterName[0:2];
    code = findCode(prov, prefix);
    parentCode = findParentCode(prov, prefix);

    db = MySQLdb.connect("localhost", "root", "magic123", "bitnami_wordpress", charset='utf8',
                         unix_socket='/opt/wordpress-4.9.8-0/mysql/tmp/mysql.sock');
    cursor = db.cursor();
    cursor.execute("SET NAMES utf8");

    ret = insertDataIntoTermRelationShip(db,cursor, postId, str(code));
    if ret == 0 :
        if parentCode!= -1 :
            insertDataIntoTermRelationShip(db, cursor, postId, str(parentCode));

    db.close();
    print("insertProvDataIntoDB end..");

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

def findCode(prov,prefix):
    for item in prov:
        if item.has_key(1) :
            if prefix == item[1] :
                return item[4];
    return 89;

def findParentCode(prov,prefix):
    parentStr ="";
    for item in prov:
        if item.has_key(1) :
            if prefix == item[1] :
                if item.has_key(3):
                    parentStr = item[3];
                    break;
                else:
                    return -1;
    if len(parentStr) == 0 :
        return -1;
    return findCode(prov, parentStr);

def readProvince(path):
    list=[];
    workbook = xlrd.open_workbook(path);
    for sheetTemp in workbook.sheets() :
        rowNum = sheetTemp.nrows;
        if rowNum != 0:
            #print(rowNum);
            list = readProvinceSheet(sheetTemp);
            break;
    return list;

def readProvinceSheet(sheetSrc):
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
            #print(rowList);
            result.append(rowList);
    return result;

def parseProvince(list):
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
    try:
        os.remove(r"""/opt/tmp/post_id_from_xml.txt""");
    except OSError as tmp:
        print(tmp);

    province = readProvince(r"""/home/jankin/examples/province_pinyin_1006.xlsx""");
    if province == None :
        print("province is not exist.")
        exit(1);

    prov = None;
    if len(province)>0:
        prov = parseProvince(province);

    if prov == None :
        print("prov is None");
        exit(2);

    result=[];
    result = readXml(r"""/home/jankin/examples/xml/1/1.xml""");
    createAllPost(result,prov);
    print("end.");