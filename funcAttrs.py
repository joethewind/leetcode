def foo():
    return True

def bar():
    'bar() does not do much'
    return True

foo.__doc__ = 'foo does not do much'

foo.tester = '''
if foo():
    print ('passed')
else:
    print('failed')
'''

for eachAttr in dir():
    obj = eval(eachAttr)
    if isinstance(obj,type(foo())):
        if hasattr(obj,'__doc__'):
            print('\nfunction"%s"has a doc string:\n\t%s'%(eachAttr,obj.__doc__))
        if hasattr(obj,'tester'):
            print('function "%s"has a tester...executing'%eachAttr )
            exec(obj.tester)
        else:
            print('function "%s" has no tester..skiping'%eachAttr)
    else:
        print('"%s" is not a function'%eachAttr)

