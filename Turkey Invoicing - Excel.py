import time
import numpy as np
import pandas as pd
import win32com.client
from datetime import datetime
import timeit
import subprocess
import json
import psutil

print("hi1")

# dir_1 = 'C:\\Turkey_Invoicing\\Raw_Files\\'
# dir_2 = 'C:\\Turkey_Invoicing\\Updated_Files\\'
var_2 = (datetime.now()).strftime("%d.%m.%Y")

with open('Cred\\Details.json') as f:
    contents = json.load(f)

# Carrier File
start = timeit.default_timer()

data = pd.read_excel('Raw_Files\\' + 'S5 Deliveries ' + var_2[0:5] + ' BB release.xlsx', header=None)

df_Carrier = data.copy()

column_names = ['Customer', 'PID No.', 'HPON', 'Exact Delivery Date', 'E-Delivery']
df_Carrier.columns = column_names

df_Carrier['Exact Delivery Date'] = df_Carrier['Exact Delivery Date'].dt.date

# Saving all the unique orders from "S5 Deliveries" file to a ".txt" file
df_Orders = pd.read_excel('Raw_Files\\' + 'S5 Deliveries ' + var_2[0:5] + ' BB release.xlsx', header=None)
list_Orders = df_Orders[2].unique()


def write_list_to_file(file_path, input_list):
    with open(file_path, 'w') as file:
        for item in input_list:
            file.write(str(item) + '\n')


file_path = 'Updated_Files\\Orders_S4.txt'

# Write the list contents to the file
write_list_to_file(file_path, list_Orders)

# Check for ARUBA Products
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

# Navigating to sales order transaction
session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ_OM_E2E_DB"
session.findById("wnd[0]/tbar[0]/btn[0]").press()

# Going to "Turkey" variant
session.findById("wnd[0]/tbar[1]/btn[25]").press()                                    # Clearing all the defaults
session.findById("wnd[0]/usr/ctxtP_VARI").text = "/TURKEY_BB"

# Taking orders from the ".txt" file
session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[23]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = contents['Directory'] + "Updated_Files\\"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "Orders_S4.txt"

session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]").sendVKey(8)

# Saving the S4 Excel file in "Raw Files" folder
session.findById("wnd[0]/tbar[1]/btn[18]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = contents['Directory'] + 'Raw_Files\\'
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "S4 Data " + var_2 + '.xlsx'
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Going back to the main screen
session.findById("wnd[0]/tbar[0]/btn[15]").press()
session.findById("wnd[0]/tbar[0]/btn[15]").press()

# Navigate to sales order transaction
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
session.findById("wnd[0]/tbar[0]/btn[0]").press()

time.sleep(3)

list_aruba_orders = []

start_index = 0

while start_index < len(df_Carrier['HPON'].unique()):

    try:

        current_index = start_index

        # Enter sales order number
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = df_Carrier['HPON'].unique()[start_index]
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus()
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()

        if session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/txtTVKBT-BEZEI").text == 'Aruba':
            list_aruba_orders.append(df_Carrier['HPON'].unique()[start_index])

            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()

        else:
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()

        time.sleep(1)

        start_index = start_index + 1

    except Exception as e:

        print(f"There was an error for order {df_Carrier['HPON'].unique()[start_index]}")

        start_index = start_index + 1
        list_orders_check = []

        list_orders_check.append(df_Carrier['HPON'].unique()[start_index - 1])  # -1 to get the order which gave error

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

df_hybrid = df_Carrier.copy()

# Not having "ARUBA" orders & "Check" (Orders that threw an error in S4) orders
try:
    df_hybrid = df_hybrid[(~df_hybrid['HPON'].isin(list_aruba_orders)) & (~df_hybrid['HPON'].isin(list_orders_check))]

except:
    df_hybrid = df_hybrid[(~df_hybrid['HPON'].isin(list_aruba_orders))]

df_hybrid = df_hybrid.reset_index()
df_hybrid = df_hybrid.drop('index', axis=1)

df_hybrid.to_excel('Updated_Files\\' + 'Carrier_Final.xlsx', sheet_name='HYBRID IT', index=False)

df_Aruba = df_Carrier[df_Carrier['HPON'].isin(list_aruba_orders)]  # Having "ARUBA" orders
df_Aruba = df_Aruba.reset_index()
df_Aruba = df_Aruba.drop('index', axis=1)

with pd.ExcelWriter('Updated_Files\\' + 'Carrier_Final.xlsx', mode='a', engine='openpyxl') as writer:
    df_Aruba.to_excel(writer, sheet_name='ARUBA', index=False)

try:
    if len(list_orders_check) == 0:
        orders_check_df = pd.DataFrame(columns=['Customer', 'PID No.', 'HPON', 'Exact Delivery Date',
                                                'E-Delivery'])
        with pd.ExcelWriter('Updated_Files\\' + 'Carrier_Final.xlsx', mode='a', engine='openpyxl') as writer:
            orders_check_df.to_excel(writer, sheet_name='Fallout', index=False)

    else:
        orders_check_df = df_Carrier[df_Carrier['HPON'].isin(list_orders_check)]
        with pd.ExcelWriter('Updated_Files\\' + 'Carrier_Final.xlsx', mode='a', engine='openpyxl') as writer:
            orders_check_df.to_excel(writer, sheet_name='Fallout', index=False)

except:
    print("list_orders_check is not there")

# Identifying Products from which BB will be removed / Merge Data Frame (Getting products from PIDs)
df1 = pd.read_excel('Updated_Files\\' + 'Carrier_Final.xlsx', sheet_name='HYBRID IT')

df_S4 = pd.read_excel('Raw_Files\\' + "S4 Data " + var_2 + '.xlsx')  # Reading S4 Data

df_S4 = df_S4.rename(columns={'SO': 'HPON', 'Shipment ID (Pack ID)': 'PID No.'})
df_S4 = df_S4[['HPON', 'PID No.', 'SOI', 'HLI', 'Product No', 'Product Desc', 'ExtIStat', 'HSTAT', 'ISTAT', 'Item Billing Block']]

# merged_df_all = This Data frame contains all the orders, i.e., "DELIVERED" as well orders that are not "DELIVERED"

merged_df_all = pd.merge(df1, df_S4, on=['HPON', 'PID No.'], how='left')
merged_df_all['Order Found'] = ~merged_df_all['Product No'].isnull()

for i in range(0, len(merged_df_all['Exact Delivery Date'])):
    merged_df_all.loc[i, 'Exact Delivery Date'] = (merged_df_all.loc[i, 'Exact Delivery Date']).date()

# Checking if we have any order that is not "DELIVERED" (Filtering out "DELIVERED" orders)

merged_df = merged_df_all.copy()

# Removing orders for which few items are in "TRUE" status and few in "FALSE"
list_orders_remove = merged_df_all[merged_df_all['Order Found'] == False]['HPON'].unique()
merged_df = merged_df[~(merged_df['HPON'].isin(list_orders_remove))]

merged_df = merged_df[(merged_df['ExtIStat'] == 'DELIVERED')]  # merged_df data frame will contain only 'DELIVERED' orders

if len(merged_df[(merged_df['ExtIStat'] == 'DELIVERED')]) != 0:

    merged_df.to_excel('Updated_Files\\' + 'Merged_Data.xlsx', sheet_name="Merged", index=False)

    # Adding "BB_Place sheet", containing Orders (Products) from which BB needs to be removed

    df_True_Values = merged_df[merged_df['Order Found'] == True]
    df_False_Values = merged_df[merged_df['Order Found'] == False]

    df_True_Values = df_True_Values.sort_values(by=['HPON', 'SOI'])  # Line Items should be ordered
    df_True_Values.reset_index(inplace=True)
    df_True_Values.drop('index', axis=1, inplace=True)

    # Filtering out only "SERVICE" products
    list_service = ['SERVICE', 'Service', 'service', 'SVC', 'Svc', 'svc']
    df_service_products = df_S4[df_S4['Product Desc'].apply(lambda x: any(substring in x for substring in list_service))]
    df_service_products = df_service_products[(df_service_products['HPON'].isin(df_True_Values['HPON'])) & (df_service_products['ExtIStat'] == 'DELIVERED')]

    df_service_products.reset_index(inplace=True)
    df_service_products = df_service_products.drop('index', axis=1)

    df_True_Values = pd.concat([df_True_Values, df_service_products])

    df_True_Values = df_True_Values.drop_duplicates(subset=['HPON', 'SOI'])

    df_True_Values = df_True_Values.sort_values(by=['HPON', 'SOI'])    # Line Items should be ordered
    df_True_Values.reset_index(inplace=True)
    df_True_Values.drop('index', axis=1, inplace=True)

    df_True_Values = df_True_Values.fillna('None')


    df_support = pd.DataFrame()
    for i in range(0, len(df_True_Values)):
        df_S4_2 = df_S4[(df_S4['HPON'] == df_True_Values['HPON'][i]) & (df_S4['HLI'] == df_True_Values['SOI'][i])]

        if len(df_S4_2) == 0:
            continue
        else:
            df_support = pd.concat([df_support, df_S4_2])

    df_True_Values = pd.concat([df_True_Values, df_support])

    # df_True_Values = df_True_Values.drop_duplicates(subset=['Product No'])
    df_True_Values = df_True_Values.drop_duplicates(subset=['HPON', 'SOI'])

    df_True_Values = df_True_Values.sort_values(by=['HPON', 'SOI'])  # Line Items should be ordered
    df_True_Values.reset_index(inplace=True)
    df_True_Values.drop('index', axis=1, inplace=True)

    df_True_Values = df_True_Values.fillna('None')


    # Function to fill the "E-Delivery" numbers for "MCMB, MGBH" Products


    def filling_na(Column_Name):
        for i in range(0, len(df_True_Values[Column_Name])):
            if df_True_Values[Column_Name][i] != 'None':
                continue
            else:
                # df_True_Values[Column_Name][i] = df_True_Values[Column_Name][i-1]
                df_True_Values.loc[i, Column_Name] = df_True_Values.loc[i - 1, Column_Name]

    filling_na('E-Delivery')  # Filling E-Delivery numbers

    # Comparing S4 data frame (df_S4) and merged_data frame (merged_df)
    df_Place_BB = pd.DataFrame()

    for hpon in df_True_Values['HPON'].unique():
        merged_sois = df_True_Values.loc[df_True_Values['HPON'] == hpon, 'SOI'].unique()
        df_sois = df_S4.loc[df_S4['HPON'] == hpon, 'SOI'].unique()
        new_sois = np.setdiff1d(df_sois, merged_sois)
        new_rows = pd.DataFrame({'HPON': hpon, 'SOI': new_sois})
        df_Place_BB = pd.concat([df_Place_BB, new_rows], ignore_index=True)

    df_Place_BB = df_Place_BB[['HPON', 'SOI']].sort_values(by=['HPON', 'SOI']).reset_index(drop=True)

    # Removing "INVOICED" (INV) items from this sheet, as BB is already removed so need to place BB again
    df_Place_BB = df_Place_BB.merge(df_S4[['HPON', 'SOI', 'ISTAT', 'Item Billing Block']], how='left', on=['HPON', 'SOI'])
    df_Place_BB = df_Place_BB[~(df_Place_BB['ISTAT'] == 'INV')]
    df_Place_BB = df_Place_BB[~((df_Place_BB['Item Billing Block'] == 'J1') | (df_Place_BB['Item Billing Block'] == '1A'))]

    df_Place_BB.reset_index(inplace=True)
    df_Place_BB.drop('index', axis=1, inplace=True)

    # Filter df_S4 to only include orders with all products in "DELIVERED" status
    df_S4_filtered = df_S4.groupby("HPON").filter(lambda x: (x["ExtIStat"] == "DELIVERED").all())

    # This is the data frame (result) which contains orders from which we have to remove BB from all the products

    # Merge df_S4_filtered with merged_df on HPON
    result = pd.merge(df_S4_filtered[["HPON"]], df_True_Values, on="HPON", how="inner")
    result = result.drop_duplicates(subset=["HPON"])
    result.reset_index(inplace=True)
    result.drop('index', axis=1, inplace=True)
    result = result['HPON']

    if len(result) != 0:
        # Getting the E-Delivery numbers for all the orders in the "result" data frame

        # Creating an empty list to store the results
        output_list = []

        # looping through each order in the result dataframe
        for order in result:

            # getting the E-Delivery values for the order from the df_True_Values dataframe
            edelivery_values = df_True_Values[df_True_Values['HPON'] == order]['E-Delivery'].tolist()

            # appending the order and E-Delivery values to the output list
            for edelivery_value in edelivery_values:
                output_list.append({'HPON': order, 'E-Delivery': edelivery_value})

        # creating a new dataframe from the output list
        output_df = pd.DataFrame(output_list)

        output_df = output_df.drop_duplicates(subset=['HPON', 'E-Delivery'])
        output_df = output_df[~(output_df['E-Delivery'] == 'None')]
        output_df = output_df.reset_index()
        output_df.drop('index', axis=1, inplace=True)

        # Saving all the three sheets to "Merged_Data" excel

        df_True_Values = df_True_Values[~df_True_Values['HPON'].isin(output_df['HPON'])]
        with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
            df_True_Values.to_excel(writer, sheet_name='BB_Remove', index=False)

        df_Place_BB = df_Place_BB[~df_Place_BB['HPON'].isin(output_df['HPON'])]
        with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
            df_Place_BB.to_excel(writer, sheet_name='BB_Place', index=False)

        with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name='BB_Remove_All', index=False)

    else:

        # Saving all the three sheets to "Merged_Data" excel
        with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
            df_True_Values.to_excel(writer, sheet_name='BB_Remove', index=False)

        with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
            df_Place_BB.to_excel(writer, sheet_name='BB_Place', index=False)

        output_df = pd.DataFrame(columns=['HPON', 'E-Delivery'])  # Creating a blank data frame, as "output_df" is empty
        with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name='BB_Remove_All', index=False)


else:

    df_merged_all = pd.DataFrame(columns=['Customer', 'PID No.', 'HPON', 'Exact Delivery Date', 'E-Delivery',
                                          'SOI', 'HLI', 'Product No', 'Product Desc', 'ExtIStat', 'HSTAT', 'ISTAT',
                                          'Order Found'])
    df_merged_all.to_excel('Updated_Files\\' + 'Merged_Data.xlsx', sheet_name="Merged", index=False)

    df_True_Values = pd.DataFrame(columns=['Customer', 'PID No.', 'HPON', 'Exact Delivery Date', 'E-Delivery',
                                           'SOI', 'HLI', 'Product No', 'Product Desc', 'ExtIStat', 'HSTAT', 'ISTAT',
                                           'Order Found'])
    with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
        df_True_Values.to_excel(writer, sheet_name='BB_Remove', index=False)

    df_Place_BB = pd.DataFrame(columns=['Customer', 'PID No.', 'HPON', 'Exact Delivery Date', 'E-Delivery',
                                        'SOI', 'HLI', 'Product No', 'Product Desc', 'ExtIStat', 'HSTAT', 'ISTAT',
                                        'Order Found'])
    with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
        df_Place_BB.to_excel(writer, sheet_name='BB_Place', index=False)

    output_df = pd.DataFrame(columns=['HPON', 'E-Delivery'])  # Creating a blank data frame
    with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
        output_df.to_excel(writer, sheet_name='BB_Remove_All', index=False)

list_HPON_fallout = merged_df_all[(merged_df_all['HLI'] == 0) & (merged_df_all['ExtIStat'] != 'DELIVERED')]['HPON']
fallout_df = merged_df_all[(merged_df_all['HPON'].isin(list_HPON_fallout))]
fallout_df = fallout_df.drop_duplicates('HPON')
fallout_df.reset_index(inplace=True)
fallout_df = fallout_df.drop('index', axis=1)

# Creating a new DataFrame from the list with the same column name as the target column
new_rows_df = pd.DataFrame({'HPON': list_orders_remove})

# Concatenating the new DataFrame with the original DataFrame
fallout_df = pd.concat([fallout_df, new_rows_df], ignore_index=True)

if len(fallout_df) == 0:
    fallout_df['HPON'] = pd.DataFrame(columns=['HPON'])

    with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
        fallout_df['HPON'].to_excel(writer, sheet_name='Fallout', index=False)

else:
    with pd.ExcelWriter('Updated_Files\\' + 'Merged_Data.xlsx', mode='a', engine='openpyxl') as writer:
        fallout_df['HPON'].to_excel(writer, sheet_name='Fallout', index=False)

stop = timeit.default_timer()
print('Time: ', stop - start)
