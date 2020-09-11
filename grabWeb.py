from urllib.request import urlretrieve

def firstNonblank(lines):
    for eachline in lines:
        if not eachline.strip( ):
            continue
        else:
            return eachline

def firstLast(webpage):
    f = open(webpage,'r',encoding='UTF-8')
    lines = f.readlines()
    f.close()
    print(firstNonblank(lines))
    lines.reverse()
    print(firstNonblank(lines))

def download(url='http://www',process=firstLast):
    try:
        retrval = urlretrieve(url)[0]
    except IOError:
        retrval = None
    if retrval: #do some processing
        process(retrval)

if __name__=='__main__':
    download()
        