from enum import Enum


class Part:

    part_no = None
    short_desc = None
    long_desc = None
    price = None

    def __init__(self, _part_no, _short_desc, _long_desc, _price):
        self.part_no = _part_no
        self.short_desc = _short_desc
        self.long_desc = _long_desc
        self.price = _price

    
    def print(self):
        string = "Part No: {}, Short Desc: {}, Long Desc: {}, Price: {}".format(self.part_no,self.short_desc,self.long_desc,self.price)
        print(string)

    def toString(self):
        return 'Part No: {}\nShort Description: {}\nLong Description: {}\nPrice: {}'.format(self.part_no, self.short_desc, self.long_desc, self.price)
        
    def getPrice(self):
        return self.price

    def getShortDesc(self):
        return self.short_desc

    def getPartNo(self):
        return self.part_no
    
    def getLongDesc(self):
        return self.long_desc
    
class Switch(Part):

    NoOfPort = None

    def __init__(self, _part_no, _short_desc, _long_desc, _price, _no_of_port = None, _no_of_uplink = None, _wire_speed = None, _switch_throughput = None):
        super().__init__(_part_no, _short_desc, _long_desc, _price)
        self.no_of_port = _no_of_port
        self.no_of_uplink = _no_of_uplink
        self.wire_speed = _wire_speed
        self.switch_throughput = _switch_throughput
    
    def getNoOFPort(self):
        return self.no_of_port
    
    def getNoOfUplink(self):
        return self.no_of_uplink
    
    def getWireSpeed(self):
        return self.wire_speed
    
    def getThroughput(self):
        return self.switch_throughput

    def print(self):
        super().print()
        print("No of Port: {}".format(self.no_of_port))


class Optics(Part):

    range = None
    bandwidth = None

    def __init__(self, _part_no, _short_desc, _long_desc, _price, _range = None, _bandwidth = None):
        super().__init__(_part_no, _short_desc, _long_desc, _price)
        self.range = _range
        self.bandwidth = _bandwidth

    def getRange(self):
        return self.range

    def getBandwidth(self):
        return self.bandwidth

    def print(self):
        super().print()
        print("Range: {}".format(self.range))

    

class IMM(Part):

    def __init__(self, _part_no, _short_desc, _long_desc, _price):
        super().__init__(_part_no, _short_desc, _long_desc, _price)

class IOM(Part):

    def __init__(self, _part_no, _short_desc, _long_desc, _price):
        super().__init__(_part_no, _short_desc, _long_desc, _price)
    
class MDA(Part):

    def __init__(self, _part_no, _short_desc, _long_desc, _price):
        super().__init__(_part_no, _short_desc, _long_desc, _price)

class OS(Part):

    def __init__(self, _part_no, _short_desc, _long_desc, _price):
        super().__init__(_part_no, _short_desc, _long_desc, _price)




    