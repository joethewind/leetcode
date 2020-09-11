myTuple = ['123','2','4']

i = iter(myTuple)

print(i.__next__())


fetch = iter(myTuple)
while True:
    try:
        i = fetch.__next__()
    except StopIteration:
        break

[(x+1,y+1) for x in range(3) for y in range(5)]

