from enum import Enum


class Part:

    partNo = None
    short_desc = None
    long_desc = None
    price = None

    def __init__(self, partNo, short_desc, long_desc, price):
        self.partNo = partNo
        self.short_desc = short_desc
        self.long_desc = long_desc
        self.price = price

    
    def print(self):
        string = "Part No: {}, Short Desc: {}, Long Desc: {}, Price: {}".format(self.partNo,self.short_desc,self.long_desc,self.price)
        print(string)

    def toString(self):
        return 'Part No: {}\nShort Description: {}\nLong Description: {}\nPrice: {}'.format(self.partNo, self.short_desc, self.long_desc, self.price)
        
    def getPrice(self):
        return self.price

    def getShortDesc(self):
        return self.short_desc

    def getPartNo(self):
        return self.partNo
    
    def getLongDesc(self):
        return self.long_desc
    
class Switch(Part):

    NoOfPort = None

    def __init__(self,partNo, short_desc, long_desc, price, NoOfPort):
        super().__init__(partNo, short_desc, long_desc, price)
        self.NoOfPort = NoOfPort

    def print(self):
        super().print()
        print("No of Port: {}".format(self.NoOfPort))


class Optics(Part):

    range = None

    def __init__(self,partNo, short_desc, long_desc, price, range):
        super().__init__(partNo, short_desc, long_desc, price)
        self.range = range

    def print(self):
        super().print()
        print("Range: {}".format(self.range))



    
