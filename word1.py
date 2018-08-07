import docx;

def test1():
    print("test1 begin.");
    #print("""E:\lichuan\势坤科技\examples""");
    filePath = """E:\lichuan\势坤科技\examples\成都电子科技大学信息表.docx""";
    print(filePath);
    file = docx.Document(filePath);
    print("段落数:" + str(len(file.paragraphs)));
    print("test1 end.");


if __name__ == "__main__":
    print("word test.");
    test1();

