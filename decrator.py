from time import ctime,sleep

def tsfunc(func):
    def wrappendfunc():
        print('[%s]%s()called'%(ctime(),func.__name__))
        return func()
    return wrappendfunc()

@tsfunc
def foo():
    pass


if __name__=='__main__':
    foo()
    sleep(4)

    for i in range(2):
        sleep(1)
        foo()