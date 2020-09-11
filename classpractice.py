class Hotelrent():
    def __init__(self,rt,sales = 0.085, rm = 0.1):
        self.salesTax = sales
        self.roomTax = rm
        self.roomRate = rt
    def calcTotal(self,days = 1):
        daily =  round((self.roomRate*(1+self.salesTax+self.roomTax)),2)
        return float(days)*daily

sfo = Hotelrent(299,0.06,0.2)
print(sfo.calcTotal(2))

bj = Hotelrent(299)
print(bj.calcTotal(2))
