import PySimpleGUI as sg
import os
import copy
import colorama
from colorama import Fore
from pywinauto.application import Application
import time
import pyautogui
import pyautogui as pag
home = os.path.expanduser('~')
cwd2 = os.path.join(home, 'Documents')
cwd3 = os.path.join(cwd2, 'Inventory Deluxe')
cwd4 = os.path.join(cwd3, 'SAP GUI')
cwd5 = os.path.join(cwd3, 'SAP Inventory')

def getinventoryinfo(returns):
    """
    Function to go through each text document to extract the data for our relevant areas
    """
    PNlist = []
    with open(cwd5 + r'\Inventory_Template.csv','r') as f:
        data = []
        for line in f:
            data_line = line.rstrip().split(',')
            data.append(data_line)
    pndesc = [x[0:2] for x in data]
    
#_____________________________________________________________________________________
#Get rid of description
    
    for elem in range(len(data)):
        data[elem].pop(1)
#_____________________________________________________________________________________
#Get rid of top line with table heaers
    
    data.pop(0)
#_____________________________________________________________________________________
#If not a digit, remove it for formatting
    
    for elem in data:
        if elem[0].isdigit():
            pass
        else:
            data.pop(data.index(elem))
    
    for line in sorted(data):
        PNlist.append(line[0])
    for elem in data:
        for item in range(elem.count('')):
            elem.remove('')
    
    if returns == 'PNlist':
        return PNlist
    if returns == 'data':
        return data
    if returns == 'pndesc':
        return pndesc

def performonSAP():
    """
    Function to do all the work on SAP
    """
    home = os.path.expanduser('~')
    PNlist = getinventoryinfo('PNlist')
    priddir = ''  
    for elem in home:
        priddir += elem
    prid = priddir[9:16]
    
    directory = "SAP"
    parent_dir = 'C:\\Users\\' + prid + "\\OneDrive - AZCollaboration\\Documents\\"
    path = os.path.join(parent_dir, directory)
    pathgui = os.path.join(path, "SAP GUI")
    filelist = [f for f in os.listdir(pathgui)]
    
    if os.path.isdir(path) == False: #If SAP folder does not exist, make it
        os.mkdir(path)
    else:
        if os.path.isdir(pathgui) == True: #If SAP GUI folder exists, clear it to prevent errors saving to text from SAP
            for g in filelist:
                os.remove(os.path.join(pathgui,g))
    count = 0
    for item in PNlist:
        if count == 0:
            time.sleep(3)
        else:
            time.sleep(1)
            
        if count >= 1:
            pag.typewrite(str(item)); pag.hotkey('enter')
            time.sleep(0.5)
            pag.hotkey('f9')
            time.sleep(0.6)
            pag.hotkey('down')
            pag.hotkey('enter')
            time.sleep(0.6)
            pag.typewrite(str(item)+".txt")
            
            for i in range(6):
                pag.hotkey('tab')
            pag.typewrite(pathgui)
            time.sleep(0.5)
            pag.hotkey('enter')
            time.sleep(3)
            pag.hotkey('f3')
            count += 1

        elif count == 0:
            pag.typewrite(str(item)); pag.hotkey('tab')
            pag.typewrite('5120'); pag.hotkey('tab')
            pag.typewrite('0010'); 
            
            for i in range(13):
                pag.hotkey('tab')
            pag.typewrite('GP1')
            pag.hotkey('enter')
            time.sleep(0.5)
            pag.hotkey('f9')
            time.sleep(0.6)
            pag.hotkey('down')
            pag.hotkey('enter')
            time.sleep(0.6)
            pag.typewrite(str(item)+".txt")
            
            for i in range(6):
                pag.hotkey('tab')
            pag.typewrite(pathgui)
            time.sleep(0.5)
            pag.hotkey('enter')
            time.sleep(3)
            pag.hotkey('f3')
            count += 1

def getSAPinfo(area):
    """
    Function to get info from SAP text files created in performonSAP()
    """
    priddir = ''  
    for elem in home:
        priddir += elem
    prid = priddir[9:16]
    pnlist = sorted(getinventoryinfo('PNlist'))
    datalist = sorted(getinventoryinfo('data'))
    mList = []
    gList = []
    bList = []
    tableofdata = []
    directory = "SAP"
    parent_dir = 'C:\\Users\\' + prid + "\\OneDrive - AZCollaboration\\Documents\\"
    path = os.path.join(parent_dir, directory)
    pathgui = os.path.join(path, "SAP GUI")
    sapinfo = [[]] * len(pnlist)
    sapinfo3 = []
    
    for file in os.listdir(pathgui):
        count = 0
        tableofdata = []
        p = os.path.join(pathgui, file)
        if os.path.isfile(p):
            with open(p) as f:
                for line in f:
                    tableofdata.append(line.split())
            
            indices = [index for index, element in enumerate(tableofdata) if element == []] 
            mindex = [index for index, element in enumerate(tableofdata) if element == ['101','Media','Prep']] 
            bindex = [index for index, element in enumerate(tableofdata) if element == ['102','Buffer','Prep']] 
            gwindex = [index for index, element in enumerate(tableofdata) if element == ['015','CIP','storage']] 

            if area == 'Media':
                for elem in mindex:
                    indices.insert(0,elem)
                        
            elif area == 'Buffer':
                for elem in bindex:
                    indices.insert(0,elem)
                    
            elif area == 'GW':
                for elem in gwindex:
                    indices.insert(0,elem)   
                    
            indices2 = sorted(indices)

            if area == 'Media':
                if mindex:
                    indexofMedia = indices2.index(mindex[0])
                    mLold = tableofdata[indices2[indexofMedia]:indices2[indexofMedia+1]]
                    mL = [x for xs in mLold for x in xs]
                    for i in range(5):
                        mL.pop(0)
                    mL.pop(2)
                    mL.pop(2)
                    if len(mL) >= 3:
                        mL.pop(-1)
                        mL.pop(-1)
                        mL.pop(2)
                        mL.pop(2)
                    mL.insert(0,tableofdata[1][1])
                    mList.append(mL)
            
            if area == 'Buffer':
                if bindex:
                    indexofBuffer = indices2.index(bindex[0])
                    bLold = tableofdata[indices2[indexofBuffer]:indices2[indexofBuffer+1]]
                    bL = [x for xs in bLold for x in xs]
                    for i in range(5):
                        bL.pop(0)
                    bL.pop(2)
                    bL.pop(2)
                    if len(bL) >= 3:
                        bL.pop(-1)
                        bL.pop(-1)
                        bL.pop(2)
                        bL.pop(2)
                    bL.insert(0,tableofdata[1][1])
                    bList.append(bL)
                    
            if area == 'GW':
                if gwindex:
                    indexofGW = indices2.index(gwindex[0])
                    gLold = tableofdata[indices2[indexofGW]:indices2[indexofGW+1]]
                    gL = [x for xs in gLold for x in xs]
                    for i in range(5):
                        gL.pop(0)
                    gL.pop(2)
                    gL.pop(2)
                    if len(gL) >= 3:
                        gL.pop(-1)
                        gL.pop(-1)
                        gL.pop(2)
                        gL.pop(2)
                    gL.insert(0,tableofdata[1][1])
                    gList.append(gL)
                    
    sapinfo2 = [[el] for el in pnlist] # Convert PN list into list of lists to append to
    
    #Add to our list of sap item info for each area 
    if area == 'Media':
        for col in mList:
            for col2 in sapinfo2:
                if col[0] == col2[0]:
                    for item in col:
                        col2.append(item)
    if area == 'Buffer':
        for col in bList:
            for col2 in sapinfo2:
                if col[0] == col2[0]:
                    for item in col:
                        col2.append(item)      
    if area == 'GW':
        for col in gList:
            for col2 in sapinfo2:
                if col[0] == col2[0]:
                    for item in col:
                        col2.append(item)

    #Remove duplicate entries
    for elem in sapinfo2:
        if len(elem) > 1:
            elem.pop(0)
    
    for row in datalist:
        if '' in row:
            for i in range(row.count('')):
                row.remove('')
                
    # Modify SAP values to reflect number of containers/parts
    for pn in sapinfo2:
        if pn[0] == '10099':
            for i in range(2,len(pn),2):
                pn[i] = str(int(float(pn[i])/3.78))
                
        if pn[0] == '4000249' or pn[0] == '4000250' or pn[0] == '4102555' or pn[0] == '4102546':
            for i in range(2,len(pn),2):
                pn[i] = str(int(int(pn[i])/100))
        
        if pn[0] == '4000163':
            for i in range(2,len(pn),2):
                pn[i] = str(int(int(pn[i])/250))
        
        if pn[0] == '4000074' or pn[0] == '4102545':
            for i in range(2,len(pn),2):
                pn[i] = str(int(int(pn[i])/50))
        
        if pn[0] == '6000371':
            for i in range(2,len(pn),2):
                pn[i] = str(int(int(pn[i])/500))
    
        
        if pn[0] == '4000078' or pn[0] == '4102567' or pn[0] == '4000151':
            for i in range(2,len(pn),2):
                pn[i] = str(int(int(pn[i])/25))
        
        
        if pn[0] == '4100891' or pn[0] == '4100893' or pn[0] == '4100892' or pn[0] == '4102371':
            for i in range(2,len(pn),2):
                pn[i] = str(int(int(pn[i])/10))
                
    sapinfo3 = copy.deepcopy(sapinfo2) # Create deepcopy of all the SAP info to not modify the original list
    
    def comparedata(sapinfo3, datalist):
        """
        Compare SAP to inventory
        """
        contactmmlist = []
        consumeitalllist = []
        consumelist = []
        sapdict = {}
        userdict = {}
        dict_3 = {}
        
        # Remove PNs to turn into dict of {BN:quantities}
        for item in sapinfo3:
            item.pop(0)
        
        for item in datalist:
            item.pop(0)
            
        # Flatten out the nested list
        sapinfo4 = [x for x in sapinfo3 for x in x]
        userinfo4 = [x for x in datalist for x in x]
        
        # Turn string quantities into integers
        for item in range(1,len(userinfo4),2):
            userinfo4[item] = int(userinfo4[item])*-1
        for item in range(1,len(sapinfo4),2):
            sapinfo4[item] = int(sapinfo4[item])
            
        # Populate dict
        for item in range(0,len(sapinfo4),2):
            sapdict[sapinfo4[item]] = sapinfo4[item+1]
        for item in range(0,len(userinfo4),2):
            userdict[userinfo4[item]] = userinfo4[item+1]
        
        # Add SAP and user dicts                            
        dict_3 = {**sapdict, **userdict}
        for key, value in dict_3.items():
            if key in sapdict and key in userdict:
                dict_3[key] = value + sapdict[key]                        
            
        return dict_3
        
        
    dict_3 = comparedata(sapinfo3, datalist)
    dict_3_action = [key for key, value in dict_3.items() if value != 0] # Get all PNs that must be consumed/contacted
   
    listtocheck = []
    for item in sapinfo2:
        for elem in dict_3_action:
            if elem in item:
                listtocheck.append([item[0],elem,dict_3[elem]])
                
    # Make new lists to filter what needs to be consumed or sent to MM
    consumelist = []
    contactmmlist = []
    for item in listtocheck:
        if item[2] > 0:
            consumelist.append(item)
        elif item[2] < 0:
            contactmmlist.append(item)
    
    return consumelist,contactmmlist


headings = ['PN', 'BN','Quantity']
sg.theme('DarkBrown1')

layout = [
         [sg.Text('Area'),sg.Text(' ',size=(5,1)),sg.Checkbox('Media',key = 'mediakey'),sg.Checkbox('Buffer', key='bufferkey'),sg.Checkbox('GW',key='gwkey')],
    [sg.Text('')],
         [sg.Button('Extract SAP info', key=lambda: performonSAP())],
         [sg.Button('Calculate inventory', key=lambda: getSAPinfo(area))],[sg.Text('')],
         [sg.Text('Consume'),sg.Text('',size=(29,1)),sg.Text('Contact MM')],
         [sg.Multiline('',size=(25,5),key='consumekey'),sg.Text(' ',size=(9,1)),sg.Multiline('',size=(25,5),key='contactkey')]]
         
layout2 = [[sg.Text('1. Select an area checkbox (only one may be selected)')],
           
           [sg.Text('NOTE: SAP must be open. SAP must be on the LS26 screen. Ensure the highlighted text box is "Material"')],
           [sg.Text('2. Press "Extract SAP info" and open the SAP application, ensuring above guidelines are met')],
           [sg.Text('NOTE: You will have a few seconds to get to the LS26 screen before the application starts woring')],
           [sg.Text('NOTE: Do not move your mouse or touch your keyboard while the application is working')],
           [sg.Text('3. After step 2 is complete, return to this application and press "Calculate inventory"')],
           [sg.Text('')],
           [sg.Text('NOTE: The folder containing the consumable inventory MUST be..')],
           [sg.Text(cwd5+"\\Inventory_Template.xlsx")]]

grp = [[sg.TabGroup([[sg.Tab('Auto Rec',layout),sg.Tab('Instructions',layout2)]])]]
window = sg.Window('Auto Rec',grp)
while True:             # Event Loop`
    event, values = window.read() 
    print(event, values)       
    if event == sg.WIN_CLOSED or event == 'Exit':
        break   
    if values['mediakey'] == True:
        area = 'Media'
    elif values['bufferkey'] == True:
        area = 'Buffer'
    elif values['gwkey'] == True:
        area = 'GW'

    if callable(event):
        event()
        
        window['consumekey'].update(getSAPinfo(area)[0])
        window['contactkey'].update(getSAPinfo(area)[1])
window.close()