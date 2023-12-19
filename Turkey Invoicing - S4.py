import time
import numpy as np
import pandas as pd
import win32com.client
from datetime import datetime
import timeit
import subprocess
import json
import psutil

# dir_1 = 'C:\\Turkey_Invoicing\\Raw_Files\\'
# dir_2 = 'C:\\Turkey_Invoicing\\Updated_Files\\'
var_2 = (datetime.now()).strftime("%d.%m.%Y")

with open('Cred\\Details.json') as f:
    contents = json.load(f)

# S4 Part
# Fixed
path = r'C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe'
subprocess.Popen(path)

time.sleep(1)

# Connect to SAP GUI
SapGuiAuto = win32com.client.GetObject("SAPGUI")
if not type(SapGuiAuto) == win32com.client.CDispatch:
    raise Exception("SAP GUI not found")

application = SapGuiAuto.GetScriptingEngine
if not type(application) == win32com.client.CDispatch:
    raise Exception("SAP GUI scripting engine not found")

# Connect to SAP system
connection = application.OpenConnection(contents['DCPConnectionName'], True)
if not type(connection) == win32com.client.CDispatch:
    raise Exception("SAP system not found")

session = connection.Children(0)
if not type(session) == win32com.client.CDispatch:
    raise Exception("SAP session not found")

# Set up event handling
if hasattr(session, "On"):
    session.On("ScriptingCommandBarPanel", session.Events)

# Login to SAP system
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = contents['user']
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = contents['password']
session.findById("wnd[0]").sendVKey(0)

# Navigate to sales order transaction
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA02"
session.findById("wnd[0]/tbar[0]/btn[0]").press()

# Main Part
# Removing BB from orders in "BB_Remove" sheet
df = pd.read_excel('Updated_Files\\' + 'Merged_Data.xlsx', sheet_name='BB_Remove')
dict_BB_Remove = {hpon_value: df.loc[df['HPON'] == hpon_value, 'SOI'].tolist() for hpon_value in df['HPON'].unique()}
List_Order_Number = [i for i in dict_BB_Remove.keys()]


def Func_BB_Remove(Order_Number):
    # Enter sales order number
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = Order_Number
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus()
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)

    # Checking If Header Level BB is there or not, if "YES", removing it, otherwise, no action.

    # Get the value of the text box
    textbox = session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK")
    textbox.SetFocus()

    text = textbox.Text

    # Check if the desired text is present in the copied text
    if "Billing Block" or "Invoice Hold" in text:
        print("The desired text is present in the text box - Header Level BB removed" + "\n")

        # Removing Header Level BB
        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK").key = " "

    else:
        print("No Header Level BB" + "\n")

    # Code to go into a particular product ####

    # filter the data frame to get all rows with the specified order number
    order_df = df[df['HPON'] == Order_Number]

    # iterate through all the items under the order
    List_Item_Numbers = []

    for index, row in order_df.iterrows():
        Item_Number = row['SOI']  # Column name that contains the item numbers
        List_Item_Numbers.append(str(Item_Number))

    Line_Items = List_Item_Numbers

    flag = 0
    flag2 = 0

    # Get the SAP table object
    Product_Table = session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG")

    # Setting the number of visible rows to display at once
    Visible_Rows = Product_Table.VisibleRowCount

    for item in Line_Items:
        found = False

        # Checking if product is in the initial visible rows
        if flag2 == 0:
            for row_index in range(Visible_Rows):
                cell = Product_Table.GetCell(row_index, 0)
                if cell is not None and cell.Text == item:
                    found = True
                    Product_Table.getAbsoluteRow(row_index).selected = True
                    break

        elif flag2 == 1:
            for row_index in range(Visible_Rows):
                Product_Table = session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG")
                cell = Product_Table.GetCell(row_index, 0)
                if cell is not None and cell.Text == item:
                    found = True
                    row_index = row_index + flag
                    Product_Table.getAbsoluteRow(row_index).selected = True
                    break

        # If not found in initial visible rows, scrolling down and checking again
        while not found:
            session.findById("wnd[0]").sendVKey(82)  # Scroll down
            flag = flag + Visible_Rows
            flag2 = 1
            for row_index in range(Visible_Rows):
                Product_Table = session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG")
                cell = Product_Table.GetCell(row_index, 0)
                if cell is not None and cell.Text == item:
                    found = True
                    row_index = row_index + flag
                    Product_Table.getAbsoluteRow(row_index).selected = True
                    break

    # Remove Item Level Billing Block

    session.findById("wnd[0]/mbar/menu[1]/menu[1]/menu[2]").select()
    session.findById("wnd[1]/tbar[0]/btn[7]").press()

    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11").select()

    # Updating E-Delivery Number
    # filter the data frame to get all rows with the specified order number
    order_df = df[df['HPON'] == Order_Number]

    # iterate through all the products under the order
    List_E_Del = []

    for index, row in order_df.iterrows():
        E_Del_Number = row['E-Delivery']  # Column name that contains the product name
        List_E_Del.append(E_Del_Number)

    for i in set(List_E_Del):
        if i != "None":
            # select the ZTDN item in the table and double click it
            shell = session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell")
            shell.selectItem("ZTDN", "Column1")
            shell.ensureVisibleHorizontalItem("ZTDN", "Column1")
            shell.topNode = "ZES9"
            shell.doubleClickItem("ZTDN", "Column1")

            # Updating the next text
            existingText = session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text

            appendedText = str(i)

            if existingText == '\r':
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = appendedText
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                    8, 8)

            else:
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = existingText + "\n" + appendedText
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                    8, 8)
        else:
            continue

    session.findById("wnd[0]").sendVKey(11)  # Saving the order

    # To handle the prompt which comes after entering the E-Delivery Number and saving it
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").Press()

    except:
        # Handle any exceptions that occur during execution
        print("No prompt")


# Calling "Func_BB_Remove" Function
if len(df) != 0:
    start = timeit.default_timer()
    for i in List_Order_Number:
        Func_BB_Remove(i)
    stop = timeit.default_timer()
    print('Time: ', stop - start)
else:
    pass

# Placing BB on orders from "BB_Place" sheet
df = pd.read_excel('Updated_Files\\' + 'Merged_Data.xlsx', sheet_name='BB_Place')
dict_BB_Place = {hpon_value: df.loc[df['HPON'] == hpon_value, 'SOI'].tolist() for hpon_value in df['HPON'].unique()}
List_Order_Number = [i for i in dict_BB_Place.keys()]


def Func_BB_Place(Order_Number):
    # Enter sales order number
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = Order_Number  # Dummy Case
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus()
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)

    # Code to go into a particular product
    # filter the data frame to get all rows with the specified order number
    order_df = df[df['HPON'] == Order_Number]

    # iterate through all the items under the order
    List_Item_Numbers = []

    for index, row in order_df.iterrows():
        Item_Number = row['SOI']  # Column name that contains the item numbers
        List_Item_Numbers.append(str(Item_Number))
    print(List_Item_Numbers)

    Line_Items = List_Item_Numbers

    flag = 0
    flag2 = 0

    # Get the SAP table object
    Product_Table = session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG")

    # Setting the number of visible rows to display at once
    Visible_Rows = Product_Table.VisibleRowCount

    for item in Line_Items:
        found = False

        # Checking if product is in the initial visible rows
        if flag2 == 0:
            for row_index in range(Visible_Rows):
                cell = Product_Table.GetCell(row_index, 0)
                if cell is not None and cell.Text == item:
                    found = True
                    Product_Table.getAbsoluteRow(row_index).selected = True
                    break

        elif flag2 == 1:
            for row_index in range(Visible_Rows):
                Product_Table = session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG")
                cell = Product_Table.GetCell(row_index, 0)
                if cell is not None and cell.Text == item:
                    found = True
                    row_index = row_index + flag
                    Product_Table.getAbsoluteRow(row_index).selected = True
                    break

        # If not found in initial visible rows, scrolling down and checking again
        while not found:
            session.findById("wnd[0]").sendVKey(82)  # Scroll down
            flag = flag + Visible_Rows
            flag2 = 1
            for row_index in range(Visible_Rows):
                Product_Table = session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG")
                cell = Product_Table.GetCell(row_index, 0)
                if cell is not None and cell.Text == item:
                    found = True
                    row_index = row_index + flag
                    Product_Table.getAbsoluteRow(row_index).selected = True
                    break

    # Adding Item Level Billing Block
    session.findById("wnd[0]/mbar/menu[1]/menu[1]/menu[2]").select()
    session.findById("wnd[1]/usr/ctxtRV45A-S_FAKSP").text = "1A"
    session.findById("wnd[1]/tbar[0]/btn[7]").press()

    session.findById("wnd[0]").sendVKey(11)

    # To handle the prompt which comes after entering the E-Delivery Number and saving it
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").Press()

    except:
        # Handle any exceptions that occur during execution
        print("No prompt")


# Calling "Func_BB_Place" Function
if len(df) != 0:
    start = timeit.default_timer()
    for i in List_Order_Number:
        Func_BB_Place(i)
    stop = timeit.default_timer()
    print('Time: ', stop - start)
else:
    pass

# Removing BB from the all products under a particular order for orders in "BB_Remove_All" sheet
df = pd.read_excel('Updated_Files\\' + 'Merged_Data.xlsx', sheet_name='BB_Remove_All')
dict_BB_Remove_All = {hpon_value: df.loc[df['HPON'] == hpon_value, 'E-Delivery'].tolist() for hpon_value in
                      df['HPON'].unique()}
List_Order_Number = [i for i in dict_BB_Remove_All.keys()]


def Func_BB_Remove_All(Order_Number):
    # Enter sales order number
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = Order_Number
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus()
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)

    # Checking If Header Level BB is there or not, if "YES", removing it, otherwise, no action.
    # Get the value of the text box
    textbox = session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK")
    textbox.SetFocus()

    text = textbox.Text

    print(Order_Number)

    # Check if the desired text is present in the copied text
    if "Billing Block" or "Invoice Hold" in text:
        print("The desired text is present in the text box - Header Level BB removed" + "\n")

        # Removing Header Level BB
        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK").key = " "

    else:
        print("No Header Level BB" + "\n")

    # Code to go into a particular product

    # filter the data frame to get all rows with the specified order number
    # order_df = df[df['HPON'] == Order_Number]

    session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_MKAL").press()

    # Remove Item Level Billing Block from all the products

    session.findById("wnd[0]/mbar/menu[1]/menu[1]/menu[2]").select()
    session.findById("wnd[1]/tbar[0]/btn[7]").press()

    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11").select()

    # Updating E-Delivery Number
    # List_E_Del = []

    # filter the data frame to get all rows with the specified order number
    order_df = df[df['HPON'] == Order_Number]

    # iterate through all the products under the order
    List_E_Del = []

    for index, row in order_df.iterrows():
        E_Del_Number = row['E-Delivery']  # Column name that contains the product name
        List_E_Del.append(E_Del_Number)

    for i in set(List_E_Del):

        if i != "None":
            # select the ZTDN item in the table and double click it
            shell = session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell")
            shell.selectItem("ZTDN", "Column1")
            shell.ensureVisibleHorizontalItem("ZTDN", "Column1")
            shell.topNode = "ZES9"
            shell.doubleClickItem("ZTDN", "Column1")

            # Updating the next text
            existingText = session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text

            appendedText = str(i)

            if existingText == '\r':
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = appendedText
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                    8, 8)

            else:
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = existingText + "\n" + appendedText
                session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                    8, 8)
        else:
            print("Mission Successful")
            continue

    session.findById("wnd[0]").sendVKey(11)

    # To handle the prompt which comes after entering the E-Delivery Number and saving it
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").Press()

    except:
        # Handle any exceptions that occur during execution
        print("No prompt")


# Calling "Func_BB_Remove" Function
if len(df) != 0:
    start = timeit.default_timer()
    for i in List_Order_Number:
        Func_BB_Remove_All(i)
    stop = timeit.default_timer()
    print('Time: ', stop - start)
else:
    pass

# Closing the SAP Window
session.findById("wnd[0]").Close()
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

# Defining the SAP Logon process name
process_name = 'saplogon.exe'

# Find the process ID (PID) of the SAP Logon process
for proc in psutil.process_iter(['pid', 'name']):
    if proc.info['name'] == process_name:
        saplogon_pid = proc.info['pid']
        break
else:
    saplogon_pid = None

# Terminating the SAP Logon process if found
if saplogon_pid:
    process = psutil.Process(saplogon_pid)
    process.terminate()
    process.wait()
    print("SAP Logon window closed.")
else:
    print("SAP Logon process not found.")
