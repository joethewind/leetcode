def counter(start):
    count = [start]

    def incr():
        count[0] += 1
        return count[0]

    return incr

count = counter(5)

print(count())