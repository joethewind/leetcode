from operator import add,mul
from functools import partial,reduce

a = list()

add1 = partial(add,1) #add1(x) = add(x) + 1
mul100 = partial(mul,100) #mul100(x) = mul(100,x)
base2bad = partial(int,base=16)

a = map(lambda x:x+2,range(6))
b = map(lambda x,y:(x+y,x-y),[1,2,4],[2,5,7])
c = map(None,[1,3,5],[2,4,6])
d = base2bad('1F')

for i in b:
    print(i)

print('total num is ', reduce((lambda x,y:x+y),range(5)))
print(d)