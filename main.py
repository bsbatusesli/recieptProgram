import PySimpleGUI as sg
from Part import *
from Bundle import *
import pandas as pd
import os
import pickle


windowSize = (1000 ,600)

partList = []
bundle = Bundle()

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

def loadall(filename):
    with open(filename, "rb") as f:
        while True:
            try:
                yield pickle.load(f)
            except EOFError:
                break

## Load previously added parts
def loadAllParts(filename, df):

    with open(filename, 'rb') as input:
        partListCode = pickle.load(input)

    for partNo in partListCode:
        partList.append(createPart(partNo, df))

    partListName = updatePartListName(partList)



## Create the window
def createWindow():

    
    mainLayout =[           [sg.Button("Create Part", key = '-GO_PART-')],
                            [sg.Button("Create Bundle", key = '-GO_BUNDLE-')],
                            [sg.Button("Display Current Bundle", key = '-DISPLAY_BUNDLE-')],
                            [sg.Button("Exit")] ]

    createPartLayout = [    [sg.Text("Create Part")], 
                            [sg.Text('Enter Item Code: '), sg.InputText(key = '-ITEMCODE-')],
                            [sg.Button("ADD", key = '-ADDPART-'), sg.Button("Cancel", key = '-EXITPART-')],
                            [sg.Text(size = (100,10), key = '-ConfirmationMessage-')]   ]

    createBundleLayout =[   [sg.Text('Select Part')],
                            [sg.Listbox(values = partList, select_mode='extended', key='listbox', size=(100, 6))], 
                            [sg.Text('Quantity: '), sg.InputText(key = '-ITEMQUANTITY-', size = (10,5))],
                            [sg.Text('Total Price: ', key = '-TOTALPRICE-', size = (50,1))],
                            [sg.Button("Add", key = '-ADDBUNDLE-'),sg.Button("Go Main Page", key = '-EXITBUNDLE-')],
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
                            [sg.Button("Go Main Page", key = '-EXITACTIVE-'),sg.Button("Create Excel File", key = '-CREATEEXCEL-'), sg.Button("Save Bundle", key = '-SAVEBUNDLE-')],[sg.Input(key = '-BUNDLE_PATH-' ), sg.FileBrowse("Browse"), sg.Button("Load Bundle", key = '-LOADBUNDLE-')], 
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
            bundle_data = bundle.toDataFrame()
            window['-BUNDLETABLE-'].update(values = bundle_data)
            window['-TOTALPRICE2-'].update('Total Price: ' + str(round(bundle.calculateTotalPrice(), 2)) + ' €')

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

        # Adding part to bundle
        if event == '-ADDBUNDLE-' :
            selection = values['listbox'][0]

            #catch quantity type exceptions
            try:
                quantity = int(values['-ITEMQUANTITY-'])
                index = int(partListName.index(selection))
                message_str = 'Added item: {}\n Quantity: {}'.format(selection,quantity)

                if quantity > 0:
                    window['-Message-'].update('ADDED ITEM to Bundle\n' + message_str)
                    bundle.addPart(partList[index], quantity)
                    window['-TOTALPRICE-'].update('Total Price: ' + str(round(bundle.calculateTotalPrice(), 2)) + ' €')
            except: 
                window['-Message-'].update('Please Enter Valid Quantity')

        # create excel file
        if event == '-CREATEEXCEL-':
            ####  !!!!!!!!!!!!      Note: Cancel doesn't work       !!!!!!!!!  ####
            filename = sg.popup_get_text('Save Sheet\nEnter Excel File Name')
            pd.DataFrame(data = bundle_data, columns = bundle_headings).to_excel("{}.xlsx".format(str(filename)))

        # save bundle
        if event == '-SAVEBUNDLE-':
            filename = sg.popup_get_text('Save Bundle\nEnter Bundle Name')
            save_object(bundle, 'bundle_{}.pkl'.format(str(filename)))
            bundle.print()

        '''if event == '-LOADBUNDLE-':
            bundle_path = str(values['-BUNDLE_PATH-']).split(sep = "/")
            file_name = bundle_path[len(bundle_path)-1]
            print(file_name)
            with open(file_name, 'rb') as input:
                loaded_bundle = pickle.load(input)
                print(loaded_bundle.toString())
            print(loaded_bundle.toString())
            bundle_data = loaded_bundle.toDataFrame()
            window['-BUNDLETABLE-'].update(values = bundle_data)
            window['-TOTALPRICE2-'].update('Total Price: ' + str(round(bundle.calculateTotalPrice(), 2)) + ' €')'''
            

        #------ INTERNAL BUTTON FUNCTIONS ENDS------------#
            
    window.close()



if __name__ == '__main__':
    main()


## TODO LIST:
# Save bundle with name
# excel name popup optimization . Cancel doesnt work
# Added part classification
# delete part
# delete part from bundle


'''## ---------TEST OBJECTS ---------##
product1 = createPart('3HE10492AB', df)
partList.append(product1)
product2 = createPart('3HE10503AA', df)
partList.append(product2)

bundle.addPart(product1,1)
bundle.addPart(product2,2)
## ---------TEST OBJECTS ---------##'''