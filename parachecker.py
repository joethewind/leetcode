import stack

def parchecker(symbolString):
    s = stack.Stack()
    balacned = True
    index = 0

    while index < len(symbolString) and balacned:
            symbol = symbolString[index]
            if symbol == '(':
                s.push(symbol)
            else:
                if s.isEmpty():
                    balacned = False
                else:
                    s.pop()

            index = index + 1

    if s.isEmpty() and balacned == True:
        return True
    else:
        return False

print(parchecker('()'))

