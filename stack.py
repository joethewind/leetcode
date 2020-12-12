class Stack:
    def __init__(self):
        self.item = []

    def isEmpty(self):
        return self.item == []

    def push(self):
        self.item.append()

    def pop(self):
        return self.item.pop()

    def size(self):
        return len(self.item)