#!/user/bin/env python
# coding:UTF-8
# func_closure: 这个属性仅当函数是一个闭包时有效
# 指向一个保存了所引用到的外部函数的变量cell的元组
# 如果该函数不是一个内部函数，则始终为None。这个属性也是只读的。
output = '<int %r id = %#0x val =%d>'
w = x = y = z = 1


def f1():
    print('f1() start')
    x = y = z = 2

    def f2():
        print('f2() start')
        y = z = 3

        def f3():
            print('f3() start')
            z = 4
            print(output % ('w', id(w), w))
            print(output % ('x', id(x), x))
            print(output % ('y', id(y), y))
            print(output % ('z', id(z), z))

        print('f3.func_closure')
        clo = f3.__closure__
        if clo:
            print("f3 closure vars:", [str(c) for c in clo])
        else:
            print('no f3 closure vars')
            print('f2() call f3()')
        print('in f2 call f3')
        f3()
        print('f2() end')
    print('f1.func_closure')
    clo = f2.__closure__
    if clo:
        print("f2 closure vars:", [str(c) for c in clo])
    else:
        print("no f2 closure vars")
    print('in f1() call f2()')
    f2()
    print('f1() end')
print('f1.func_closure')
clo = f1.__closure__
if clo:
    print("f1 closure vars:", [str(c) for c in clo])
else:
    print("no f1 closure vars")
print('call f1()')
f1()
