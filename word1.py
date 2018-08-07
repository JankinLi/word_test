import docx;

def test1():
    print("test1 begin.");
    #print("""E:\lichuan\势坤科技\examples""");
    filePath = """E:\lichuan\势坤科技\examples\成都电子科技大学信息表.docx""";
    print(filePath);
    file = docx.Document(filePath);
    print("段落数:" + str(len(file.paragraphs)));
    # 输出每一段的内容
    for para in file.paragraphs:
        print(para.text)

    # 输出段落编号及段落内容
    for i in range(len(file.paragraphs)):
        print("第" + str(i) + "段的内容是：" + file.paragraphs[i].text)
    print("test1 end.");


if __name__ == "__main__":
    print("word test.");
    test1();
	print("end.");

