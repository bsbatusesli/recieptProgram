from src.Part import *

class Bundle:


    def __init__(self):
        self.connected_parts = []
        self.total_price = 0
        self.name = None

    def __init__(self, name):
        self.name = name
        self.connected_parts = []
        self.total_price = 0


    def setName(self,name):
        self.name = name

    def getName(self):
        return self.name
    

    def addPart(self, Part, quantity):
        itemCodes = []

        for parts in self.connected_parts:
            itemCodes.append(parts[0].getPartNo())
        
        #check part no exist in the bundle
        if Part.getPartNo() in itemCodes:
            index = itemCodes.index(Part.getPartNo())
            self.connected_parts[index][1] += quantity 
        else:
            self.connected_parts.append([Part,quantity])

    def removePart(self, Part, quantity):
        for parts in self.connected_parts:
            if parts[0] is Part:
                if parts[1] > quantity:
                    parts[1] -= quantity
                    return
                if parts[1] == quantity:
                    self.connected_parts.remove(parts)
                    return
                if parts[1] < quantity:
                    raise InsufficientQuantity

        ## if part is not in the connected part
        raise PartDoesNotExist

    def print(self):
        for parts in self.connected_parts:
            parts[0].print()
            print("Quantity: {}".format(parts[1]))

    def toString(self):
        output_str = ''
        for parts in self.connected_parts:
            output_str +='{}\t\t {} €\t\t {}\n'.format(parts[0].short_desc,parts[0].price, parts[1])
        return output_str

    def toDataFrame(self):
        displayList = []
        for parts in self.connected_parts:
            displayList.append([parts[0].short_desc,parts[0].price, parts[1]])
        return displayList

    def calculateTotalPrice(self):
        sum = 0
        for parts in self.connected_parts:
            sum = (parts[0].getPrice() * parts[1]) + sum
        self.total_price = sum
        return sum

    def getConnectedParts(self):
        return self.connected_parts
    
    def getTotalPrice(self):
        return self.total_price
    
    def clearBundle(self):
        self.connected_parts = []
        self.total_price = 0
        
class InsufficientQuantity(Exception):
    pass

class PartDoesNotExist(Exception):
    pass







