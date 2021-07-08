import PySimpleGUI as sg
from Part import *
from Bundle import *
import pandas as pd
import os
import pickle

class UnpicklingError(Exception):
    pass

## get values from excel by looking at itemcode
def createPart(itemCode, df):
    return Part(itemCode, df.loc[itemCode]['Short Description'], df.loc[itemCode]['Long Description'], df.loc[itemCode]['Product Market Price in EUR'])

## create partlist name array (UPDATE BEFORE USE in every time)
def updatePartListName(partList):
    partListName = []
    for i in range (len(partList)):
        partListName.append(partList[i].getShortDesc())
    return partListName

def updateBundleListName(bundleList):
    bundleListName = []
    for i in range (len(bundleList)):
        bundleListName.append(bundleList[i].getName())
    return bundleListName

def updatePartListCode(partList):
    partListCode = []
    for i in range (len(partList)):
        partListCode.append(partList[i].getPartNo())
    return partListCode

def isPartExists(partList, itemCode):
    for part in partList:
        if part.getPartNo() == itemCode:
            return True
    return False

## Save objects in pkl file
def save_object(obj, filename):
    with open(filename, 'wb') as output:  # Overwrites any existing file.
        pickle.dump(obj, output, pickle.HIGHEST_PROTOCOL)

## Load previously added parts
def loadAllParts(filename, df):

    with open(filename, 'rb') as input:
        partListCode = pickle.load(input)

    for partNo in partListCode:
        partList.append(createPart(partNo, df))

    partListName = updatePartListName(partList)

def checkFileFormat(filename, fileType):
    try:
        if filename.split('.')[1] == fileType:
            return True
        else:
            return False
    except:
        return False


## Returns list contains itemCode and quantity
def packBundle(bundle_toPack):
    packed_bundle = []
    for parts in bundle_toPack.connected_parts:
        packed_bundle.append([parts[0].partNo, parts[1]])
    return packed_bundle

def unpackBundle(bundle_toUnpack, bundleName):

    bundleList.append(Bundle(bundleName))
    bundleListName = updateBundleListName(bundleList)
    bundleIndex = len(bundleList) - 1 
    partListCode = updatePartListCode(partList)
    for part in bundle_toUnpack:
        if part[0] in partListCode:
            index = partListCode.index(part[0])
            bundleList[bundleIndex].addPart(partList[index], part[1])
        else:
            ## TODO: Test it
            product = createPart(part[0], df)
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

def calculateTotalPrice(reciept):
    price = 0
    for bundles in reciept:
        price += bundles[0].calculateTotalPrice() * bundles[1]
    return price

## Create the window
def createWindow():

    
    mainLayout =[           [sg.Button("Edit Parts", key = '-GO_PART-', size = (8,8), font = 'Arial 10 bold')],
                            [sg.Button("Edit Bundle", key = '-GO_BUNDLE-', size = (8,8), font = 'Arial 10 bold')],
                            [sg.Button("Edit Reciept", key = '-GO_RECIEPT-', size = (8,8), font = 'Arial 10 bold')],
                            [sg.Button("Exit", size = (8,8), font = 'Arial 10 bold')] ]

    editPartLayout = [    [sg.Text("Edit Parts")], 
                            [sg.Text('Enter Item Code: '), sg.InputText(key = '-ITEMCODE-')],
                            [sg.Button("Add", key = '-ADDPART-'),sg.Button("Remove", key = '-REMOVEPART-'), sg.Button("Cancel", key = '-EXITPART-')],
                            [sg.Text(size = (100,10), key = '-ConfirmationMessage-')]   ]

    editBundleLayout =[  
                            [sg.Text('Active Bundle:'),sg.Text(key = '-ACTIVE_BUNDLE_NAME-' , size = (10,1))],
                            [sg.Text('Change Active Bundle'),sg.Combo(values = [], key = '-ACTIVE_BUNDLE_LIST-', enable_events = True , size = (60,1)),sg.Button("Create New", key = '-CREATEBUNDLE-')],
                            [sg.Input(key = '-BUNDLE_PATH-' ), sg.FileBrowse("Browse", file_types = (('Pickle', '*.pkl'))), sg.Button("Load Bundle", key = '-LOADBUNDLE-')], 
                            [sg.Listbox(values = partList, select_mode='extended', key='-PARTLIST-', size=(100, 6))], 
                            [sg.Text('Quantity: '), sg.InputText(key = '-ITEMQUANTITY-', size = (10,5))],
                            [sg.Button("Add", key = '-ADDBUNDLE-'),sg.Button("Remove", key = '-REMOVEBUNDLE-'),sg.Button("Go Main Page", key = '-EXITBUNDLE-')],
                            [sg.Text(size = (100,5), key = '-Message-')],
                            [sg.Table(values = bundle_data, headings=bundle_headings, max_col_width=1000, background_color='black', auto_size_columns=True,
                                        display_row_numbers=False, justification='left', num_rows=8, alternating_row_color='black', key='-BUNDLETABLE-', row_height=25)],
                            [sg.Button("Save Bundle", key = '-SAVEBUNDLE-'), sg.Button("Clear", key = '-CLEARBUNDLE-'), sg.Text('Total Price: ', key = '-TOTALPRICE_BUNDLE-', size = (50,1), font = 'Arial 20 bold')]]

    recieptLayout = [       [sg.Listbox(values = bundleList, select_mode='extended', key='-BUNDLELIST-', size=(100, 6))], 
                            [sg.Text('Quantity: '), sg.InputText(key = '-BUNDLEQUANTITY-', size = (10,5))],
                            [sg.Button("Add", key = '-ADDRECIEPT-'),sg.Button("Remove", key = '-REMOVERECIEPT-')],
                            [sg.Table(values = bundle_data, headings=bundle_headings, max_col_width=1000, background_color='black', auto_size_columns=True,
                                      display_row_numbers=False, justification='left', num_rows=8, alternating_row_color='black', key='-RECIEPTTABLE-', row_height=25)],
                            [sg.Button("Export Excel", key = '-CREATEEXCEL-'),sg.Button("Go Main Page", key = '-EXITACTIVE-'),sg.Text('Total Price: ', key = '-TOTALPRICE_RECIEPT-', size = (50,1), font = 'Arial 20 bold')]]

    layout = [[sg.Column(mainLayout, key='-MAIN-'), sg.Column(editPartLayout, visible=False, key='-PART-'),
            sg.Column(editBundleLayout, visible=False, key='-BUNDLE-'), sg.Column(recieptLayout,visible=False, key='-RECIEPT-')]]
    
    return sg.Window(title = 'SİSTAŞ PROJECT', layout = layout, size = windowSize)




## Popup function with custom layout. Return event and values
def customPopup(title, popup_layout, text = None):

    if popup_layout == 'OK_CANCEL':
        window = sg.Window(title, layout = [   [sg.Input(key = '-TEXT-')],[sg.Button("OK", key = '-OK-'), sg.Button("Cancel", key = '-CANCEL-')]])
    if popup_layout == 'ERROR_MESSAGE':
        window = sg.Window(title, layout = [   [sg.Text(text, size = (50,1))],[sg.Button("OK", key = '-OK-')]])
    event, values = window.read()
    window.close()
    return event, values

def main():

    # ----------- Reading Data From Excel ----------- #
    data = pd.read_excel(r'/Users/batuhansesli/Documents/SİSTAŞ STAJ/Product Catalogue_IP Routing_112020.xlsm', sheet_name = 'PPL', usecols = "C,E,F,K", header = 7)
    df = pd.DataFrame(data, columns = ['Item Code', 'Short Description', 'Long Description', 'Product Market Price in EUR'])
    df.set_index("Item Code", inplace = True)

    loadAllParts('saved_item_codes.pkl', df)

    window = createWindow()

    activeBundle = 0

    while True:

        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break

        elif (event == 'Exit' or event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
            if sg.popup_yes_no('Do you want to save changes?') == 'Yes':
                partListCode = updatePartListCode(partList)
                save_object(partListCode, 'saved_item_codes.pkl')
                for bundles in bundleList:
                    save_object(packBundle(bundles), 'bundle_{}.pkl'.format(str(bundles.getName())))
                break
            else:
                break
            
        # ------ NAVIGATION FUNCTIONS ------ #

        elif event == '-GO_PART-' :
            window['-MAIN-'].update(visible=False)
            window['-PART-'].update(visible=True)

        elif event == '-GO_BUNDLE-' : 
            partListName = updatePartListName(partList)
            bundleListName = updateBundleListName(bundleList)
            window.Element('-PARTLIST-').update(values = partListName)
            window.Element('-ACTIVE_BUNDLE_LIST-').update(values = bundleListName)
            window['-MAIN-'].update(visible=False)
            window['-BUNDLE-'].update(visible=True)
        
        elif event == '-GO_RECIEPT-' :
            window['-MAIN-'].update(visible=False)
            window['-RECIEPT-'].update(visible=True)
            bundleListName = updateBundleListName(bundleList)
            window.Element('-BUNDLELIST-').update(values = bundleListName)
            

        elif event == '-EXITPART-' :
            window['-PART-'].update(visible=False)
            window['-MAIN-'].update(visible=True)

        elif event == '-EXITBUNDLE-' :
            window['-BUNDLE-'].update(visible=False)
            window['-MAIN-'].update(visible=True)

        elif event == '-EXITACTIVE-' :
            window['-RECIEPT-'].update(visible=False)
            window['-MAIN-'].update(visible=True)

        # ------ NAVIGATION FUNCTIONS ENDS ------ #

        #------ INTERNAL BUTTON FUNCTIONS ------------#
        #Adding the part 
        elif event == "-ADDPART-" :

            if isPartExists(partList, values['-ITEMCODE-']) == False :
                try:
                    product = createPart(values['-ITEMCODE-'], df)
                    window['-ConfirmationMessage-'].update('ADDED ITEM\n' + product.toString())
                    partList.append(product)
                    partListName = updatePartListName(partList)
                except:
                    window['-ConfirmationMessage-'].update('Part could not found!')
            else :
                window['-ConfirmationMessage-'].update('Part is already added! ')
        
        #Removing part from Part List
        elif event == '-REMOVEPART-':
            if isPartExists(partList, values['-ITEMCODE-']) == True:
                for part in partList:
                    if part.getPartNo() == values['-ITEMCODE-']:
                        partList.remove(part)
            else:
                window['-ConfirmationMessage-'].update('Part is not found')

            
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
                        window['-Message-'].update('REMOVED ITEM\n' + message_str)            

                    # Adding part to bundle
                    elif event == '-ADDBUNDLE-':
                        bundleList[activeBundle].addPart(partList[index], quantity)
                        message_str = 'Added item: {}\n Quantity: {}'.format(selection,quantity)
                        window['-Message-'].update('ADDED ITEM\n' + message_str)
                        
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
                pd.DataFrame(data = bundle_data, columns = bundle_headings).to_excel("{}.xlsx".format(str(popup_values['-TEXT-'])))

        # Save bundle
        elif event == '-SAVEBUNDLE-':
            popup_event, popup_values = customPopup('Save Bundle','OK_CANCEL')
            if popup_event == '-OK-' and popup_values['-TEXT-'] is not '':
                save_object(packBundle(bundleList[activeBundle]), 'bundle_{}.pkl'.format(str(popup_values['-TEXT-'])))

        # Load bundle
        elif event == '-LOADBUNDLE-':
            bundle_path = str(values['-BUNDLE_PATH-']).split(sep = "/")
            file_name = bundle_path[len(bundle_path)-1]
            try:
                if checkFileFormat(file_name, 'pkl'):
                    with open(file_name, 'rb') as input:
                        loaded_bundle = pickle.load(input)

                    newBundleName = file_name.split('_')[1].split('.')[0] #get bundle name between _ and . : Saved file format bundle_<bundleName>.pkl
                    unpackBundle(loaded_bundle,newBundleName)

                    activeBundle = len(bundleList) - 1 # set active bundle to new bundle
                    updateBundleTable(window, bundleList[activeBundle])
                    bundleListName = updateBundleListName(bundleList)
                    window.Element('-ACTIVE_BUNDLE_LIST-').update(values = bundleListName)
                    window.Element('-ACTIVE_BUNDLE_NAME-').update(str(bundleList[activeBundle].getName()))
                else :
                    customPopup('ERROR', 'ERROR_MESSAGE', 'Invalid File Format. Just use pickle files(.pkl)')

            except FileNotFoundError:
                customPopup('ERROR', 'ERROR_MESSAGE', 'File Not Found')
            
        # Clear Bundle
        elif event == '-CLEARBUNDLE-':
            bundleList[activeBundle].clearBundle()
            updateBundleTable(window, bundleList[activeBundle])

        # Create New Bundle
        elif event == '-CREATEBUNDLE-':
            popup_event, popup_values = customPopup('Create New Bundle','OK_CANCEL')
            if popup_event == '-OK-' and popup_values['-TEXT-'] is not '':
                bundleList.append(Bundle(str(popup_values['-TEXT-'])))
                bundleListName = updateBundleListName(bundleList)
                window.Element('-ACTIVE_BUNDLE_LIST-').update(values = bundleListName)
        
        # Set Active Bundle in the Bundle Edit Menu
        elif event == '-ACTIVE_BUNDLE_LIST-':
            selection = values['-ACTIVE_BUNDLE_LIST-']
            activeBundle = int(bundleListName.index(selection))
            window.Element('-ACTIVE_BUNDLE_NAME-').update(str(bundleList[activeBundle].getName()))
            updateBundleTable(window, bundleList[activeBundle])


        elif event == '-ADDRECIEPT-' or event == '-REMOVERECIEPT-':
            try:
                selection = values['-BUNDLELIST-'][0]
                quantity = int(values['-BUNDLEQUANTITY-'])
                index = int(bundleListName.index(selection))

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
                    
                    updateRecieptTable(window, reciept)
                    
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
    windowSize = (700 ,600)

    partList = [] ## Stores active parts added on the stack
    bundleList = [] ## Stores bundles
    reciept = [] ## Final reciept contains bundles.

    # Table Data
    bundle_data = [['None', 'None', 'None']]
    bundle_headings = ["\t\t       Short Description       \t\t\t\t\t\t\t\t\t", "Price\t\t\t\t", 'Quantity'] # \₺ for expanding the table(using auto_size).
    reciept_data = [['None', 'None', 'None']]

    
    main()



## TODO LIST:
# Added part classification



