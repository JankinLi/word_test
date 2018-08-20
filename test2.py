import sys,getopt;
from test1.test3 import sum;

def main(argv):
    try:
        opts, args = getopt.getopt(argv,'hi:o:',["ifile=","ofile="]);
    except getopt.GetoptError as e:
        print(e.msg);
        print('test.py -i <inputfile> -o <outputfile>');
        sys.exit(2);
    except Exception:
        print('22222');
        sys.exit(1);

    for opt,arg in opts:
        if( opt == '-h'):
            print('test.py -i <inputfile> -o <outputfile>');
        elif( opt == '-i'):
            print('kkk'+arg);
        elif(opt == '-o'):
            print('is u.'+arg)
        else:
            print('unknown');

print ("你好，世界");
if __name__ == "__main__":
    print( sys.argv[0]);
    main(sys.argv[1:]);
    sum();


