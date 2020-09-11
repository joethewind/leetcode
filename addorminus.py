#!/usr/bin/env python
from operator import add, sub,mul,truediv
from random import randint, choice
ops = {'+': add, '-':sub,'*':mul,'/':truediv}
#定义一个字典
MAXTRIES = 2
def doprob():
  op = choice('+-*/')
  #用choice从'+-'中随意选择操作符
  nums = [randint(1,10) for i in range(2)]
  #用randint(1,10)随机生成一个1到10的数,随机两次使用range(2)
  nums.sort(reverse=True)
  #按升序排序
  ans = ops[op](*nums)
  #利用函数,(*nums)则相当于将nums中的元素一次作为参数传递给add这个函数
  pr = '%d %s %d = ' % (nums[0], op, nums[1])
  oops = 0
  #oops用来计算failure测试,当三次时自动给出答案
  while True:
    try:
      if int(input(pr)) == ans:
        print('correct')
        break
      if oops == MAXTRIES:
        print('answer\n %s%d' % (pr, ans))
        break
      else:
        print('incorrect... try again')
        oops += 1
    except (KeyboardInterrupt, EOFError, ValueError):
      print('invalid ipnut... try again')
def main():
  while True:
    doprob()
    try:
      opt = input('Again? [y]').lower()
      if opt and opt[0] == 'n':
        break
    except (KeyboardInterrupt, EOFError):
      break
if __name__ == '__main__':
  main()