import PySimpleGUI as sg

windowSize = (1280 ,800)

partList = [] ## Stores active parts added on the stack
bundleList = [] ## Stores bundles
reciept = [] ## Final reciept contains bundles.

activeBundle = 0

#Login heading color '20AC8A'
#kutucuk border siyah
#font Helvetica renk siyah
#size büyük 50, küçük 20, cevaplar 25,
#arkaplan beyaz

# Table Data
bundle_data = [['None', 'None', 'None']]
bundle_headings = ["\t\t       Short Description       \t\t\t\t\t\t\t\t\t", "Price\t\t\t\t", 'Quantity'] # \₺ for expanding the table(using auto_size).
reciept_data = [['None', 'None', 'None']]

heading1_font = 'Helvetica 50 bold'
heading1_size = (17,1)
heading2_font = 'Helvetica 35 bold'
heading2_size = (10,1)
text_font = 'Helvetica 20'
text_size = (10,1)
input_size_big = (30,1)
input_size_small = (6,1)
list_size = (65,6)
navigation_button_size = (14,8)
navigation_button_font = 'Helvetica 18 bold'
addremove_button_size = (20,1)
addremove_button_font = 'Helvetica 20'
function_button_size = (16,1)
function_button_font = 'Helvetica 20'
exit_button_color = ('#FFFFFF', '#7FA99F')
message_size = (100,10)
message_font = 'Helvetica 20'
popupTextFont = 'Helvatica 15'


sg.LOOK_AND_FEEL_TABLE['myTheme'] = {'BACKGROUND': '#F0F0F0',
                                        'TEXT': '#000000',
                                        'INPUT': '#FFFFFF',
                                        'TEXT_INPUT': '#000000',
                                        'SCROLL': '#99CC99',
                                        'BUTTON': ('#FFFFFF', '#20AC8A'),
                                        'PROGRESS': ('#D1826B', '#CC8019'),
                                        'BORDER': 1, 'SLIDER_DEPTH': 0, 
                                        'PROGRESS_DEPTH': 0, }

sg.theme('myTheme')

mainLayout =[           [sg.Text('', size = (1,20))],
                        [sg.Button("EDIT PARTS", key = '-GO_PART-', size = navigation_button_size, font = navigation_button_font),
                        sg.Button("EDIT BUNDLES", key = '-GO_BUNDLE-', size = navigation_button_size, font = navigation_button_font),
                        sg.Button("EDIT OFFER", key = '-GO_RECIEPT-', size = navigation_button_size, font = navigation_button_font)],
                        [sg.Text('', size = (1,20))],
                        [sg.Text('', size = (150,1)),sg.Button("Exit", size = function_button_size, font = function_button_font, button_color = exit_button_color)] ]

editPartLayout = [      
                        [sg.Text("EDIT PARTS", font = heading2_font, size = heading1_size, justification = 'center')], 
                        [sg.Text('Enter Item Code: ', font = text_font, size = heading1_size), sg.InputText(key = '-ITEMCODE-', font = text_font, size = input_size_big), sg.Button('Search', key = '-SEARCHPART-', size = (9,1), font = function_button_font)],
                        [sg.Text('Part: ', font = text_font, size = text_size), sg.Text(key = '-SEARCHEDPART-', font = text_font, size = input_size_big)],
                        [sg.Text(key = '-IN1_T-', visible = False, font = text_font, size = heading1_size)],[sg.InputText(key = '-IN1-', visible = False, font = text_font, size = input_size_small)],
                        [sg.Text(key = '-IN2_T-', visible = False, font = text_font, size = heading1_size)],[sg.InputText(key = '-IN2-', visible = False, font = text_font, size = input_size_small)], 
                        [sg.Text(key = '-IN3_T-', visible = False, font = text_font, size = heading1_size)],[sg.InputText(key = '-IN3-', visible = False, font = text_font, size = input_size_small)],
                        [sg.Button("Add", key = '-ADDPART-', size = addremove_button_size, font = addremove_button_font),sg.Text('', font = text_font, size = (text_size[0]+input_size_small[0],1)),sg.Button("Remove", key = '-REMOVEPART-', size = addremove_button_size, font = addremove_button_font)],
                        [sg.Button("Go Main Page", key = '-EXITPART-', size = function_button_size, font = function_button_font, button_color = exit_button_color)],
                        [sg.Text(key = '-ConfirmationMessage-', size = message_size, font = message_font)]   ]

editBundleLayout =[
                        [sg.Text('Change Active Bundle', font = text_font, size = heading1_size),sg.Combo(values = [], key = '-ACTIVE_BUNDLE_LIST-', enable_events = True, font = text_font, size = input_size_big),sg.Button("Create New", key = '-CREATEBUNDLE-', size = function_button_size, font = function_button_font)],
                        [sg.Text('Select Part', font = text_font, size = text_size)], 
                        [sg.Listbox(values = partList, select_mode='extended', key='-PARTLIST-', font = text_font, size = list_size)], 
                        [sg.Button("Add", key = '-ADDBUNDLE-', size = addremove_button_size, font = addremove_button_font),sg.Text('Quantity: ', font = text_font, size = text_size), sg.InputText(key = '-ITEMQUANTITY-', font = text_font, size = input_size_small), sg.Button("Remove", key = '-REMOVEBUNDLE-', size = addremove_button_size, font = addremove_button_font)],
                        [sg.Text('')],
                        [sg.Text('Active Bundle:', font = text_font + ' bold', size = heading1_size),sg.Text(key = '-ACTIVE_BUNDLE_NAME-', font = text_font, size = text_size)],
                        [sg.Table(values = bundle_data, headings=bundle_headings, max_col_width=1000, background_color='white', auto_size_columns=True,
                                    display_row_numbers=False, justification='left', num_rows=8, key='-BUNDLETABLE-', row_height=25)],
                        [sg.Text('Total Price: ', key = '-TOTALPRICE_BUNDLE-', size = (50,1), font = text_font + ' bold')],
                        [sg.Button("Save Bundle", key = '-SAVEBUNDLE-', size = function_button_size, font = function_button_font), sg.Button("Clear", key = '-CLEARBUNDLE-', size = function_button_size, font = function_button_font)],
                        [sg.Button('Go Offer Page',key = '-GO_RECIEPT_PAGE-', size = function_button_size, font = function_button_font) ,sg.Button("Go Main Page", key = '-EXITBUNDLE-', size = function_button_size, font = function_button_font, button_color = exit_button_color)]]

recieptLayout = [       [sg.Text('BUNDLES', font = heading2_font, size = heading2_size, justification = 'center')], 
                        [sg.Listbox(values = bundleList, select_mode='extended', key='-BUNDLELIST-', font = text_font, size = list_size)], 
                        [sg.Button("Add", key = '-ADDRECIEPT-', size = addremove_button_size, font = addremove_button_font),sg.Text('Quantity: ', font = text_font, size = text_size), sg.InputText(key = '-BUNDLEQUANTITY-', font = text_font, size = input_size_small), sg.Button("Remove", key = '-REMOVERECIEPT-', size = addremove_button_size, font = addremove_button_font)],
                        [sg.Text('')], 
                        [sg.Text('RECIEPT', font = heading2_font, size = heading2_size, justification = 'center')], 
                        [sg.Table(values = bundle_data, headings=bundle_headings, max_col_width=1000, background_color='white', auto_size_columns=True,
                                    display_row_numbers=False, justification='left', num_rows=8, key='-RECIEPTTABLE-', row_height=25)],
                        [sg.Text('Total Price: ', key = '-TOTALPRICE_RECIEPT-', size = (50,1), font = text_font + ' bold')],
                        [sg.Text('Discount %: ', font = text_font, size = text_size), sg.InputText(key = '-DISCOUNT-', font = text_font, size = input_size_small),sg.Button("Export Excel", key = '-CREATEEXCEL-', size = function_button_size, font = function_button_font)],
                        [sg.Button("Go Main Page", key = '-EXITACTIVE-', size = function_button_size, font = function_button_font, button_color = exit_button_color)]]

loginLayout = [
                        [sg.Text('', size = (1,20))],
                        [sg.Text('', font = heading1_font, size = heading1_size, justification = 'center')],
                        [sg.Text('Host: ', font = text_font, size = text_size), sg.InputText(key = '-HOST-', font = text_font, size = input_size_big)],
                        [sg.Text('Database: ', font = text_font, size = text_size), sg.InputText(key = '-DATABASE-', font = text_font, size = input_size_big)],
                        [sg.Text('User: ', font = text_font, size = text_size), sg.InputText(key = '-USER-', font = text_font, size = input_size_big)],
                        [sg.Text('Password: ', font = text_font, size = text_size), sg.InputText(key = '-PASSWORD-', font = text_font, size = input_size_big)],
                        [sg.Button("LOGIN",  font = text_font,key = '-LOGINBUTTON-', size = (41,1))]
                        ]



## Create the window
def createWindow():


    layout = [[sg.Column(loginLayout, key='-LOGIN-'),sg.Column(mainLayout, visible=False, key='-MAIN-'), sg.Column(editPartLayout, visible=False, key='-PART-'),
            sg.Column(editBundleLayout, visible=False, key='-BUNDLE-'), sg.Column(recieptLayout,visible=False, key='-RECIEPT-')]]
    
    return sg.Window(title = 'SİSTAŞ PROJECT', layout = layout, size = windowSize, element_justification='c')




## Popup function with custom layout. Return event and values
def customPopup(title, popup_layout, text = None, bundleNames = None):

    if popup_layout == 'EXIT':
        window = sg.Window(title, size = (500,120), layout = [  [sg.Text('')],
                                                                [sg.Text(text, font =text_font,  size = (50,1), justification = 'center')],
                                                                [sg.Text('')], 
                                                                [sg.Button("OK", key = '-OK-',font =popupTextFont,size = (10,1)),sg.Text('', size = (38,1)), sg.Button("Cancel", font =popupTextFont, key = '-CANCEL-',size = (10,1))]])
    if popup_layout == 'OK_CANCEL':
        window = sg.Window(title, size = (500,120), layout = [  [sg.Input(key = '-TEXT-', font =text_font)],
                                                                [sg.Text('')],
                                                                [sg.Text('')],
                                                                [sg.Button("OK", key = '-OK-',font =popupTextFont,size = (10,1)),sg.Text('', size = (38,1)), sg.Button("Cancel", font =popupTextFont, key = '-CANCEL-',size = (10,1))]])
    if popup_layout == 'ERROR_MESSAGE':
        window = sg.Window(title, size = (500,150), layout = [  [sg.Text('')],
                                                                [sg.Text(text, font =text_font,  size = (50,1), justification = 'center')],
                                                                [sg.Text('')],
                                                                [sg.Text('',size = (20,1)),sg.Button("OK", key = '-OK-',font =popupTextFont,size = (20,1))]])
    if popup_layout == 'LOAD':
        window = sg.Window(title, size = (500,250), layout = [  [sg.Listbox(values = bundleNames, select_mode='extended', key='-LOADBUNDLELIST-', font = text_font, size = list_size)],
                                                                [sg.Button("OK", key = '-STARTLOADING-',font =popupTextFont,size = (10,1)),sg.Text('', size = (38,1)), sg.Button("Cancel", font =popupTextFont, key = '-CANCELLOADING-',size = (10,1))] ])
    
    event, values = window.read()
    window.close()
    return event, values
