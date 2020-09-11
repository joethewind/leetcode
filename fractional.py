def fractional(n):
    if n == 0 or n == 1:
        return n
    else:
        return (n*fractional(n-1))

a = fractional(3)
print(a)