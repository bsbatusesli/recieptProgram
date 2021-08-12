from src.Bundle import *
from src.Part import *
import pandas as pd
import os
import pickle
from db.db_handler import *
from pathlib import Path
from gui import *


class UnpicklingError(Exception):
    pass

## get values from excel by looking at itemcode
def createPart(itemCode, df):
    if df.loc[itemCode]['Short Description'].startswith('SYS'):
        return Switch(itemCode, df.loc[itemCode]['Short Description'], df.loc[itemCode]['Long Description'], df.loc[itemCode]['Product Market Price in EUR'])
    elif df.loc[itemCode]['Short Description'].startswith('SFP'):
        return Optics(itemCode, df.loc[itemCode]['Short Description'], df.loc[itemCode]['Long Description'], df.loc[itemCode]['Product Market Price in EUR'])
    else:
        return Part(itemCode, df.loc[itemCode]['Short Description'], df.loc[itemCode]['Long Description'], df.loc[itemCode]['Product Market Price in EUR'])

## create partlist name array
def updatePartListName(partList):
    partListName = []
    for i in range (len(partList)):
        if partList[i] is not None:
            partListName.append(partList[i].getShortDesc())
        else:
            partListName.append('')
    return partListName

def updateBundleListName(bundleList):
    bundleListName = []
    for i in range (1,len(bundleList)):
        bundleListName.append(bundleList[i].getName())
    return bundleListName

def updatePartListCode(partList):
    partListCode = []
    for i in range (len(partList)):
        partListCode.append(partList[i].getPartNo())
    return partListCode

def isPartExists(partList, itemCode):
    for part in partList:
        try:
            if part.getPartNo() == itemCode:
                return True
        except AttributeError: #rises in index 0
            pass
    return False

## Save objects in pkl file
def save_object(obj, file_path):
    with open(file_path, 'wb') as output:  # Overwrites any existing file.
        pickle.dump(obj, output, pickle.HIGHEST_PROTOCOL)

## Load previously added parts
def loadAllParts(file_path, df):

    with open(file_path, 'rb') as input:
        partListCode = pickle.load(input)

    for partNo in partListCode:
        partList.append(createPart(partNo, df))

    partListName = updatePartListName(partList)

def loadPartsFromDB(_db):
    savedPartCount = _db.getPartCount()

    #Extends partList for inserting parts with their id locations
    while True:
        partList.append(None)
        if len(partList) > savedPartCount:
            break
    
    records = _db.getAllSwitches()
    if records is not []:
        for row in records:
            partList[row[0]] = Switch(row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8])
    
    records = _db.getAllOptics()
    if records is not []:
        for row in records:
            partList[row[0]] = Optics(row[1], row[2], row[3], row[4], row[5], row[6])

    records = _db.getOtherParts()
    if records is not []:
        for row in records:
            partList[row[0]] = Part(row[1], row[2], row[3], row[4])

def loadBundleFromDB(_db):
    savedBundleCount = _db.getBundleCount()

    #Extends partList for inserting parts with their id locations
    while True:
        bundleList.append(None)
        if len(bundleList) > savedBundleCount:
            break
    
    records = _db.getBundleItems()
    if records is not []:
        past_bid = 0
        for row in records:

            bid = row[0]
            iid = row[2]
            quantity = row[3]
            if past_bid is not bid:
                past_bid = bid
                bundleList[bid] = Bundle(row[1])
                bundleList[bid].addPart(partList[iid],quantity)
            else:
                bundleList[bid].addPart(partList[iid],quantity)

def savePartsToDB(_db):
    savedPartCount = _db.getPartCount()
    for i in range (1,len(partList)):
        if i > savedPartCount:
            _db.insertPart(partList[i].getPartNo(),partList[i].getShortDesc(),partList[i].getLongDesc(), partList[i].getPrice())
        else:
            _db.updatePart(partList[i].getPartNo(),partList[i].getShortDesc(),partList[i].getLongDesc(), partList[i].getPrice(), i) 
        if isinstance(partList[i],Switch):
            _db.updateSwitch(partList[i].getNoOFPort(),partList[i].getNoOfUplink(),partList[i].getWireSpeed(),partList[i].getThroughput(),i)
        if isinstance(partList[i],Optics):
            _db.updateOptics(partList[i].getRange(), partList[i].getBandwidth(), i)

def saveBundlesToDB(_db):
    savedBundleCount = _db.getBundleCount()
    for i in range (1,len(bundleList)):
        if i > savedBundleCount:
            _db.insertBundle(bundleList[i],i,partList)
        else:
            _db.updateBundle(bundleList[i],i,partList)





## Returns list contains itemCode and quantity
def packBundle(bundle_toPack):
    packed_bundle = []
    for parts in bundle_toPack.connected_parts:
        packed_bundle.append([parts[0].partNo, parts[1]])
    return packed_bundle

def unpackBundle(bundle_toUnpack, bundleName, _df):

    bundleList.append(Bundle(bundleName))
    bundleListName = updateBundleListName(bundleList)
    bundleIndex = len(bundleList) - 1 
    partListCode = updatePartListCode(partList)
    for part in bundle_toUnpack:
        if part[0] in partListCode:
            index = partListCode.index(part[0])
            bundleList[bundleIndex].addPart(partList[index], part[1])
        else:
            product = createPart(part[0], _df)
            partList.append(product)
            bundleList[bundleIndex].addPart(product)
            del product


## updates bundle table
def updateBundleTable(window, bundle):
    bundle_data = bundle.toDataFrame()
    window['-BUNDLETABLE-'].update(values = bundle_data)
    return bundle_data

## updates Reciept Table
def updateRecieptTable(window, reciept):
    reciept_data = []
    for bundles in reciept:
        reciept_data.append([bundles[0].getName(), '', bundles[1]])
        reciept_data.extend(updateBundleTable(window, bundles[0]))
    window['-RECIEPTTABLE-'].update(values = reciept_data)
    return reciept_data

def calculateTotalPrice(reciept):
    price = 0
    for bundles in reciept:
        price += bundles[0].calculateTotalPrice() * bundles[1]
    return price

def processOffer(values, reciept, reciept_data):
    offer_data = reciept_data
    total_price = round(calculateTotalPrice(reciept))
    offer_data.append(['','',''])
    offer_data.append(['', 'TOTAL PRICE', total_price])
    discount = int(values['-DISCOUNT-']) / 100
    if discount > 0:
        offer_data.append(['', 'DISCOUNT', total_price * discount ])
        offer_data.append(['', 'OFFER', total_price * (1 - discount) ])
        return offer_data
    else:
        raise ValueError


def main():
    # ----------- Reading Data From Excel ----------- #
    data = pd.read_excel(r'/Users/batuhansesli/Documents/SİSTAŞ STAJ/Product Catalogue_IP Routing_112020.xlsm', sheet_name = 'PPL', usecols = "C,E,F,K", header = 7)
    df = pd.DataFrame(data, columns = ['Item Code', 'Short Description', 'Long Description', 'Product Market Price in EUR'])
    df.set_index("Item Code", inplace = True)

    #loadAllParts(r'/Users/batuhansesli/sistas_project/project/localStorage/saved_item_codes.pkl', df)

    window = createWindow()


    while True:
        
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break

        elif (event == 'Exit' or event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
            popup_event = customPopup('Save Bundle','EXIT','Do you want to save changes?')
            if popup_event[0] == '-OK-':
                savePartsToDB(db_connection)
                saveBundlesToDB(db_connection)
                break
            else:
                break
            
        # ------ NAVIGATION FUNCTIONS ------ #
        
        elif event == '-LOGINBUTTON-':
            try:
                #host= values['-HOST-']
                #database= values['-DATABASE-']
                #user= values['-USER-']
                #password= values['-PASSWORD-']

                host='localhost'
                database='nokia'
                user='root'
                password='Batuhan123'

                db_connection = db(host,database,user,password)
                loadPartsFromDB(db_connection)
                loadBundleFromDB(db_connection)   
                window['-MAIN-'].update(visible=True)
                window['-LOGIN-'].update(visible=False)             
            except:
                print('Error while Connecting to DB')
        


        elif event == '-GO_PART-' :
            window['-MAIN-'].update(visible=False)
            window['-PART-'].update(visible=True)
            window['-IN1-'].update(visible = False)
            window['-IN1_T-'].update(visible = False)
            window['-IN2-'].update(visible = False)
            window['-IN2_T-'].update(visible = False)
            window['-IN3-'].update(visible = False)
            window['-IN3_T-'].update(visible = False)
            window['-SEARCHEDPART-'].update('')
            window['-ConfirmationMessage-'].update('')

        elif event == '-GO_BUNDLE-' : 
            partListName = updatePartListName(partList)
            bundleListName = updateBundleListName(bundleList)
            window.Element('-PARTLIST-').update(values = partListName)
            window.Element('-ACTIVE_BUNDLE_LIST-').update(values = bundleListName)
            window['-MAIN-'].update(visible=False)
            window['-BUNDLE-'].update(visible=True)
        
        elif event == '-GO_RECIEPT-' or event == '-GO_RECIEPT_PAGE-':
            window['-MAIN-'].update(visible=False)
            window['-BUNDLE-'].update(visible=False)
            window['-RECIEPT-'].update(visible=True)
            bundleListName = updateBundleListName(bundleList)
            window.Element('-BUNDLELIST-').update(values = bundleListName)
            
        elif event == '-EXITPART-' or event == '-EXITBUNDLE-' or event == '-EXITACTIVE-' :
            window['-PART-'].update(visible=False)
            window['-BUNDLE-'].update(visible=False)
            window['-RECIEPT-'].update(visible=False)
            window['-MAIN-'].update(visible=True)


        # ------ NAVIGATION FUNCTIONS ENDS ------ #

        #------ INTERNAL BUTTON FUNCTIONS ------------#
        #Search the Part
        elif event == '-SEARCHPART-':
            try:
                searchedPart = df.loc[values['-ITEMCODE-']]['Long Description']
                window['-SEARCHEDPART-'].update(searchedPart)

                if searchedPart.startswith('SYS'):
                    window['-IN1_T-'].update('Number of Port: ',visible = True)
                    window['-IN1-'].update(visible = True)
                    window['-IN2_T-'].update('Number of Uplink: ',visible = True)
                    window['-IN2-'].update(visible = True)
            except KeyError:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Part Could not Found')



        #Adding the part 
        elif event == "-ADDPART-" :

            if isPartExists(partList, values['-ITEMCODE-']) == False :
                try:
                    product = createPart(values['-ITEMCODE-'], df)
                    window['-ConfirmationMessage-'].update('ADDED ITEM\n' + product.toString())
                    partList.append(product)
                    partListName = updatePartListName(partList)
                except:
                    customPopup('ERROR', 'ERROR_MESSAGE', 'Part could not found!')
            else :
                customPopup('ERROR', 'ERROR_MESSAGE', 'Part is already added! ')
        
        #Removing part from Part List
        elif event == '-REMOVEPART-':
            if isPartExists(partList, values['-ITEMCODE-']) == True:
                for part in partList:
                    try:
                        if part.getPartNo() == values['-ITEMCODE-']:
                            partList.remove(part)
                            window['-ConfirmationMessage-'].update('REMOVED ITEM\n' + part.toString())
                    except AttributeError:
                        pass
            else:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Part is not found')

            
        elif event == '-ADDBUNDLE-' or event == '-REMOVEBUNDLE-':
            try:
                selection = values['-PARTLIST-'][0]
                quantity = int(values['-ITEMQUANTITY-'])
                index = int(partListName.index(selection))

                if quantity > 0:

                    # Removing part from bundle
                    if event == '-REMOVEBUNDLE-': 
                        bundleList[activeBundle].removePart(partList[index], quantity)
                        message_str = 'Removed item: {}\n Quantity: {}'.format(selection,quantity)

                    # Adding part to bundle
                    elif event == '-ADDBUNDLE-':
                        bundleList[activeBundle].addPart(partList[index], quantity)
                        message_str = 'Added item: {}\n Quantity: {}'.format(selection,quantity)
                        
                    window['-TOTALPRICE_BUNDLE-'].update('Total Price: ' + str(round(bundleList[activeBundle].calculateTotalPrice(), 2)) + ' €')
                    updateBundleTable(window, bundleList[activeBundle])

                else:
                    customPopup('ERROR', 'ERROR_MESSAGE', 'Quantity should be greater than zero')

            except InsufficientQuantity:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Insufficient Quantity for Removing')
            except PartDoesNotExist:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Part does not exist in the bundle')
            except ValueError: 
                customPopup('ERROR', 'ERROR_MESSAGE', 'Invalid Quantity Type')
            except IndexError:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Selection Error. Please make sure select one of the item')
            except UnboundLocalError:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Invalid Bundle Selection.')
            except :
                customPopup('ERROR', 'ERROR_MESSAGE', 'Upss! Try Again')

        # create excel file
        elif event == '-CREATEEXCEL-':
            popup_event, popup_values = customPopup('Create Excel','OK_CANCEL')
            if popup_event == '-OK-' and popup_values['-TEXT-'] is not '':
                reciept_data = updateRecieptTable(window, reciept)
                try:
                    pd.DataFrame(data = processOffer(values,reciept,reciept_data), columns = bundle_headings).to_excel("{}.xlsx".format(str(popup_values['-TEXT-'])))
                except ValueError:
                    customPopup('ERROR', 'ERROR_MESSAGE', 'Invalid Discount Rate')

        # Save bundle
        elif event == '-SAVEBUNDLE-':
                
            popup_event = customPopup('Save Bundle','EXIT','Do you want to save changes?')
            if popup_event[0] == '-OK-':
                savePartsToDB(db_connection)
                saveBundlesToDB(db_connection)
            else:
                pass

        # Clear Bundle
        elif event == '-CLEARBUNDLE-':
            bundleList[activeBundle].clearBundle()
            updateBundleTable(window, bundleList[activeBundle])
            window['-TOTALPRICE_BUNDLE-'].update('Total Price: ' + str(round(bundleList[activeBundle].calculateTotalPrice(), 2)) + ' €')

        # Create New Bundle
        elif event == '-CREATEBUNDLE-':
            popup_event, popup_values = customPopup('Create New Bundle','OK_CANCEL')
            if popup_event == '-OK-' and popup_values['-TEXT-'] is not '':
                bundleList.append(Bundle(str(popup_values['-TEXT-'])))
                bundleListName = updateBundleListName(bundleList)
                window.Element('-ACTIVE_BUNDLE_LIST-').update(values = bundleListName)
                activeBundle = len(bundleList) - 1
                window.Element('-ACTIVE_BUNDLE_NAME-').update(str(bundleList[activeBundle].getName()))
                updateBundleTable(window, bundleList[activeBundle])

        
        # Set Active Bundle in the Bundle Edit Menu
        elif event == '-ACTIVE_BUNDLE_LIST-':
            selection = values['-ACTIVE_BUNDLE_LIST-']
            activeBundle = int(bundleListName.index(selection)) + 1
            window.Element('-ACTIVE_BUNDLE_NAME-').update(str(bundleList[activeBundle].getName()))
            updateBundleTable(window, bundleList[activeBundle])


        elif event == '-ADDRECIEPT-' or event == '-REMOVERECIEPT-':
            try:
                selection = values['-BUNDLELIST-'][0]
                quantity = int(values['-BUNDLEQUANTITY-'])
                index = int(bundleListName.index(selection)) + 1

                if quantity > 0:

                    # Removing bundle from reciept
                    if event == '-REMOVERECIEPT-': 
                        for bundles in reciept:
                            if bundles[0] is bundleList[index]:

                                if bundles[1] > quantity:
                                    bundles[1] -= quantity

                                elif bundles[1] == quantity:
                                    reciept.remove(bundles)

                                elif bundles[1] < quantity:
                                    raise InsufficientQuantity
                                 
                    # Adding bundle to reciept
                    elif event == '-ADDRECIEPT-':

                        isExists = False
                        for bundles in reciept:
                            if bundles[0] is bundleList[index]:
                                bundles[1] += quantity
                                isExists = True
                        
                        if isExists == False:
                            reciept.append([bundleList[index], quantity])
                    
                    reciept_data = updateRecieptTable(window, reciept)
                    
                    window['-TOTALPRICE_RECIEPT-'].update('Total Price: ' + str(round(calculateTotalPrice(reciept), 2)) + ' €')
                            
                else:
                    customPopup('ERROR', 'ERROR_MESSAGE', 'Quantity should be greater than zero')

            except InsufficientQuantity:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Insufficient Quantity for Removing')
            except ValueError: 
                customPopup('ERROR', 'ERROR_MESSAGE', 'Invalid Quantity Type')
            except IndexError:
                customPopup('ERROR', 'ERROR_MESSAGE', 'Selection Error. Please make sure select one of the item')
            except :
                customPopup('ERROR', 'ERROR_MESSAGE', 'Upss! Try Again')
            

        #------ INTERNAL BUTTON FUNCTIONS ENDS------------#
            
    window.close()

if __name__ == '__main__':
    
    main()



#TODO: get requirements and setup.py create setup.py





