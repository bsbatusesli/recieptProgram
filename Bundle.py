from Part import *

class Bundle:

    connected_parts = []
    total_price = 0

    def __init__(self):
        pass

    def addPart(self, Part, quantity):
        itemCodes = []

        # get all part codes
        for parts in self.connected_parts:
            itemCodes.append(parts[0].getPartNo())
        
        #check part no exist in the bundle
        if Part.getPartNo() in itemCodes:
            #update quantity
            index = itemCodes.index(Part.getPartNo())
            self.connected_parts[index][1] += quantity 
        else:
            self.connected_parts.append([Part,quantity])

    def removePart(self, Part):
        self.connected_parts.remove(Part)

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
        return connected_parts
    
    def getTotalPrice(self):
        return total_price
    
    def clearBundle(self):
        self.connected_parts = []
        self.total_price = 0
        








