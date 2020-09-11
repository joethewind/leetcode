def testit(func,*nkargs,**kwargs):
    try:
        retval = func(*nkargs,**kwargs)
        result = (True,retval)
    except Exception:
        result = (False)
    return result

def test():
    funcs = (int,float)
    vals = (1234,12.34,'1234','12.34')

    for eachFunc in funcs:
        print('-'*20)
        for eachVal in vals:
            retval = testit(eachFunc,eachVal)
            if retval[0]:
                print('%s(%s)='%(eachFunc.__name__,eachVal),retval[1])
            else:
                print('%s(%s) = Failed:'%(eachFunc.__name__,eachVal),retval[1])

if __name__=='__main__':
    test()