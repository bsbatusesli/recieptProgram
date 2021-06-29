import PySimpleGUI as sg
from Part import *
from Bundle import *
import pandas as pd
import os
import pickle


windowSize = (1000 ,600)

partList = [] ## Stores active parts added on the stack
bundle = Bundle() ## Active Bundle Object 

# Table Data
bundle_data = [['None', 'None', 'None']]
bundle_headings = ["\t\t       Short Description       \t\t\t\t\t\t\t\t\t", "Price\t\t\t\t", 'Quantity'] # \₺ for expanding the table(using auto_size).


## get values from excel by looking at itemcode
def createPart(itemCode, df):
    return Part(itemCode, df.loc[itemCode]['Short Description'], df.loc[itemCode]['Long Description'], df.loc[itemCode]['Product Market Price in EUR'])

## create partlist name array (UPDATE BEFORE USE in every time)
def updatePartListName(partList):
    partListName = []
    for i in range (len(partList)):
        partListName.append(partList[i].getShortDesc())
    return partListName

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


## Returns list contains itemCode and quantity
def packBundle(bundle_toPack):
    packed_bundle = []
    for parts in bundle_toPack.connected_parts:
        packed_bundle.append([parts[0].partNo, parts[1]])
    return packed_bundle

def unpackBundle(bundle_toUnpack):
    partListCode = updatePartListCode(partList)
    for part in bundle_toUnpack:
        if part[0] in partListCode:
            index = partListCode.index(part[0])
            bundle.addPart(partList[index], part[1])
        else:
            ## TODO: Test it
            product = createPart(part[0], df)
            partList.append(product)
            bundle.addPart(product)
            del product

## updates bundle table
def updateBundleTable(window, bundle):
    bundle_data = bundle.toDataFrame()
    window['-BUNDLETABLE-'].update(values = bundle_data)
    window['-TOTALPRICE2-'].update('Total Price: ' + str(round(bundle.calculateTotalPrice(), 2)) + ' €')

## Create the window
def createWindow():

    
    mainLayout =[           [sg.Button("Edit Parts", key = '-GO_PART-')],
                            [sg.Button("Edit Bundle", key = '-GO_BUNDLE-')],
                            [sg.Button("Display Current Bundle", key = '-DISPLAY_BUNDLE-')],
                            [sg.Button("Exit")] ]

    createPartLayout = [    [sg.Text("Edit Parts")], 
                            [sg.Text('Enter Item Code: '), sg.InputText(key = '-ITEMCODE-')],
                            [sg.Button("Add", key = '-ADDPART-'),sg.Button("Remove", key = '-REMOVEPART-'), sg.Button("Cancel", key = '-EXITPART-')],
                            [sg.Text(size = (100,10), key = '-ConfirmationMessage-')]   ]

    createBundleLayout =[   [sg.Text('Select Part')],
                            [sg.Listbox(values = partList, select_mode='extended', key='listbox', size=(100, 6))], 
                            [sg.Text('Quantity: '), sg.InputText(key = '-ITEMQUANTITY-', size = (10,5))],
                            [sg.Text('Total Price: ', key = '-TOTALPRICE-', size = (50,1))],
                            [sg.Button("Add", key = '-ADDBUNDLE-'),sg.Button("Remove", key = '-REMOVEBUNDLE-'),sg.Button("Go Main Page", key = '-EXITBUNDLE-')],
                            [sg.Text(size = (100,10), key = '-Message-')] ]

    bundleLayout = [        
                            [sg.Table(values = bundle_data, headings=bundle_headings, max_col_width=1000,
                                    background_color='black',
                                    auto_size_columns=True,
                                    display_row_numbers=False,
                                    justification='left',
                                    num_rows=8,
                                    alternating_row_color='black',
                                    key='-BUNDLETABLE-',
                                    row_height=25)],
                            [sg.Button("Go Main Page", key = '-EXITACTIVE-'),sg.Button("Create Excel File", key = '-CREATEEXCEL-'), sg.Button("Save Bundle", key = '-SAVEBUNDLE-'), sg.Button("Clear", key = '-CLEARBUNDLE-')],
                            [sg.Input(key = '-BUNDLE_PATH-' ), sg.FileBrowse("Browse", file_types = (('Pickle', '*.pkl'))), sg.Button("Load Bundle", key = '-LOADBUNDLE-')], 
                            [sg.Text('Total Price: ', key = '-TOTALPRICE2-', size = (50,1))]]

    layout = [[sg.Column(mainLayout, key='-MAIN-'), sg.Column(createPartLayout, visible=False, key='-PART-'),
            sg.Column(createBundleLayout, visible=False, key='-BUNDLE-'), sg.Column(bundleLayout,visible=False, key='-ACTIVEBUNDLE-')]]
    
    return sg.Window(title = 'Teklif Uygulaması', layout = layout, size = windowSize)




def main():

    # ----------- Reading Data From Excel ----------- #
    data = pd.read_excel(r'/Users/batuhansesli/Documents/SİSTAŞ STAJ/Product Catalogue_IP Routing_112020.xlsm', sheet_name = 'PPL', usecols = "C,E,F,K", header = 7)
    df = pd.DataFrame(data, columns = ['Item Code', 'Short Description', 'Long Description', 'Product Market Price in EUR'])
    df.set_index("Item Code", inplace = True)

    loadAllParts('saved_item_codes.pkl', df)

    window = createWindow()

    while True:

        event, values = window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            partListCode = updatePartListCode(partList)
            save_object(partListCode, 'saved_item_codes.pkl')
            break

        # ------ NAVIGATION FUNCTIONS ------ #
        elif event == '-GO_PART-' :
            window['-MAIN-'].update(visible=False)
            window['-PART-'].update(visible=True)

        elif event == '-GO_BUNDLE-' :
            partListName = updatePartListName(partList)
            window.Element('listbox').update(values = partListName)
            window['-MAIN-'].update(visible=False)
            window['-BUNDLE-'].update(visible=True)
        
        elif event == '-DISPLAY_BUNDLE-' :
            window['-MAIN-'].update(visible=False)
            window['-ACTIVEBUNDLE-'].update(visible=True)
            updateBundleTable(window, bundle)

        elif event == '-EXITPART-' :
            window['-PART-'].update(visible=False)
            window['-MAIN-'].update(visible=True)

        elif event == '-EXITBUNDLE-' :
            window['-BUNDLE-'].update(visible=False)
            window['-MAIN-'].update(visible=True)

        elif event == '-EXITACTIVE-' :
            window['-ACTIVEBUNDLE-'].update(visible=False)
            window['-MAIN-'].update(visible=True)

        # ------ NAVIGATION FUNCTIONS ENDS ------ #

        #------ INTERNAL BUTTON FUNCTIONS ------------#
        #Adding the part 
        if event == "-ADDPART-" :

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
        if event == '-REMOVEPART-':
            if isPartExists(partList, values['-ITEMCODE-']) == True:
                for part in partList:
                    if part.getPartNo() == values['-ITEMCODE-']:
                        partList.remove(part)
            else:
                window['-ConfirmationMessage-'].update('Part is not found')

        # Adding part to bundle
        if event == '-ADDBUNDLE-' :
            try:
                selection = values['listbox'][0]
                quantity = int(values['-ITEMQUANTITY-'])
                index = int(partListName.index(selection))
                message_str = 'Added item: {}\n Quantity: {}'.format(selection,quantity)

                if quantity > 0:
                    window['-Message-'].update('ADDED ITEM\n' + message_str)
                    bundle.addPart(partList[index], quantity)
                    window['-TOTALPRICE-'].update('Total Price: ' + str(round(bundle.calculateTotalPrice(), 2)) + ' €')
            except ValueError: 
                window['-Message-'].update('Invalid Quantity')
            except IndexError:
                window['-Message-'].update('Selection Error. Please make sure select one of the item')
            except:
                window['-Message-'].update('Upss! Try Again')

        # Removing part from bundle
        if event == '-REMOVEBUNDLE-':
            try:
                selection = values['listbox'][0]
                quantity = int(values['-ITEMQUANTITY-'])
                index = int(partListName.index(selection))
                message_str = 'Removed item: {}\n Quantity: {}'.format(selection,quantity)

                if quantity > 0:
                    window['-Message-'].update('REMOVED ITEM\n' + message_str)
                    bundle.removePart(partList[index], quantity)
                    window['-TOTALPRICE-'].update('Total Price: ' + str(round(bundle.calculateTotalPrice(), 2)) + ' €')
            except InsufficientQuantity:
                window['-Message-'].update('Insufficient Quantity for Removing')
            except ValueError: 
                window['-Message-'].update('Invalid Quantity')
            except IndexError:
                window['-Message-'].update('Selection Error. Please make sure select one of the item')
            except :
                window['-Message-'].update('Upss! Try Again')
            

        # create excel file
        if event == '-CREATEEXCEL-':
            ####  !!!!!!!!!!!!      Note: Cancel doesn't work       !!!!!!!!!  ####
            filename = sg.popup_get_text('Save Sheet\nEnter Excel File Name')
            pd.DataFrame(data = bundle_data, columns = bundle_headings).to_excel("{}.xlsx".format(str(filename)))

        # Save bundle
        if event == '-SAVEBUNDLE-':
            filename = sg.popup_get_text('Save Bundle\nEnter Bundle Name')
            save_object(packBundle(bundle), 'bundle_{}.pkl'.format(str(filename)))

        # Load bundle
        if event == '-LOADBUNDLE-':
            bundle_path = str(values['-BUNDLE_PATH-']).split(sep = "/")
            file_name = bundle_path[len(bundle_path)-1]
            
            with open(file_name, 'rb') as input:
                loaded_bundle = pickle.load(input)

            bundle.clearBundle()
            unpackBundle(loaded_bundle)
            updateBundleTable(window, bundle)
            
        # Clear Bundle
        if event == '-CLEARBUNDLE-':
            bundle.clearBundle()
            updateBundleTable(window, bundle)


        #------ INTERNAL BUTTON FUNCTIONS ENDS------------#
            
    window.close()

if __name__ == '__main__':
    main()



## TODO LIST:
# excel name popup optimization . Cancel doesnt work
# Added part classification
# delete part
# multiple bundle adding 