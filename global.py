global is_this_global
is_this_global = 'abc'


def foo():

    is_this_global = 'def'
    print(is_this_global)

foo()

print(is_this_global)

