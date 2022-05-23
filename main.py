import win32com.client

import pandas as pd

import numpy as np

 

def genTable():

    resultTable = f"""

                    <table cellspacing="0" cellpadding="2" width="" border="1">

        <tbody>

        <tr>

          <th valign="top" width=""><strong>ICF_Number</strong></th>

          <th valign="top" width=""><strong>Customer_name</strong></th>

          <th valign="top" width=""><strong>Requestoer</strong></th>

          <th valign="top" width=""><strong>USN</strong></th>

          <th valign="top" width=""><strong>Product_description</strong></th>

          <th valign="top" width=""><strong>Category_Level_1</strong></th>

          <th valign="top" width=""><strong>Category_Level_2</strong></th>

          <th valign="top" width=""><strong>Warehouse_num</strong></th>

          <th valign="top" width=""><strong>Est._Monthly_Units</strong></th>

          <th valign="top" width=""><strong>UoM</strong></th>

          <th valign="top" width=""><strong>Demand_Profile</strong></th>

          <th valign="top" width=""><strong>ICF_Cost</strong></th>

          <th valign="top" width=""><strong>ICF_Cube</strong></th>

         

          </tr>

          {genTableBody()}

        </tbody>

        </table>

       

        NAME Team

        </body>

        </html>

        """

    return resultTable

 

def genTableBody():

    n = 0

    tableBody = []

    while n < len(df.index):

        tableLine = f"""<tr>

              <td valign="top" width="">{emailTable['ICF_Number'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Customer_name'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Requestor'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['USN'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Product_Description'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Category_Level_1'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Category_Level_2'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['warehouse_num'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Estimated_Monthly_Units'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['UOM'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Demand_Profile'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Total_ICF_Cost'].iloc[n]}</td>

              <td valign="top" width="">{emailTable['Total_ICF_Cube'].iloc[n]}</td>

            </tr>"""

        tableBody.append(tableLine)

        n = n + 1

    return ''.join(tableBody)

   

def appliance_check():

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.To = f'{buyer}; {merchant}'

    mail.Subject = f"Start ICF {ICF} Under Review"

    mail.HTMLBody = f"""<!DOCTYPE html>

                    <html>

                    <body>

                    NAME currently has start {ICF} under review for an appliance check. Please review the information below and indicate whether this ICF should be approved or rejected.<br>

                <strong>Cost:  {costTotal}</strong><br>

                <strong>Cube:  {cubeTotal}</strong><br>

                {genTable()}

                </body>

                </html>"""

    mail.Display(True)

   

def whirlpool_check():

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.To = f'{buyer}; {merchant}'

    mail.Subject = f"Start ICF {ICF} Under Review"

    mail.HTMLBody = f"""<!DOCTYPE html>

                    <html>

                    <body>

                    NAME currently has start {ICF} under review for an appliance check. Please review the information below and indicate whether this ICF should be approved or rejected.<br>

                    <strong>Cost:  {costTotal}</strong><br>

                    <strong>Cube:  {cubeTotal}</strong><br>

                    {genTable()}

                    </body>

                    </html>"""

    mail.Display(True)

   

def highDollar_check():

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.To = f'{merchant}; {buyer}'

    mail.Subject = f"Start ICF {ICF} Under Review"

    mail.HTMLBody = f"""<!DOCTYPE html>

                    <html>

                    <body>

                    NAME currently has start ICF {ICF} under review for high cost. Please review the information below and indicate whether this ICF should be approved or rejected.<br>

                    <strong>Cost:  {costTotal}</strong><br>

                    <strong>Cube:  {cubeTotal}</strong><br>

                    {genTable()}

                    </body>

                    </html>"""

    mail.Display(True)

   

def highCube_check():

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.To = f'{merchant}; {buyer};'

    mail.Subject = f"Start ICF {ICF} Under Review"

    mail.HTMLBody = f"""<!DOCTYPE html>

                    <html>

                    <body>

                    NAME currently has start ICF {ICF} under review for high cube. Please review the information below and indicate whether this ICF should be approved or rejected.<br>

                    <strong>Cost:  {costTotal}</strong><br>

                    <strong>Cube:  {cubeTotal}</strong><br>

                    {genTable()}

                    </body>

                    </html>"""

    mail.Display(True)

 

def costAndCube_check():

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.To = f'{merchant}; {buyer};'

    mail.Subject = f"Start ICF {ICF} Under Review"

    mail.HTMLBody = f"""<!DOCTYPE html>

                    <html>

                    <body>

                    NAME currently has start ICF {ICF} under review for high cost and high cube. Please review the information below and indicate whether this ICF should be approved or rejected.<br>

                    <strong>Cost:  {costTotal}</strong><br>

                    <strong>Cube:  {cubeTotal}</strong><br>

                    {genTable()}

                    </body>

                    </html>"""

    mail.Display(True)

   

def testMailGen():

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.To = f'{me}'

    mail.Subject = f"Start ICF {ICF} Under Review"

    mail.HTMLBody = f"""<!DOCTYPE html>

                    <html>

                    <body>

                    NAME currently has start ICF {ICF} under review for high cost and high cube. Please review the information below and indicate whether this ICF should be approved or rejected.<br>

                    <strong>Cost:  {costTotal}</strong><br>

                    <strong>Cube:  {cubeTotal}</strong><br>

                    {genTable()}

                    </body>

                    </html>"""

    mail.Display(True)

 

reasonDict = {"['Pipe Check']":testMailGen,

    "['Capacity Check & ICF Cost greater than 30000']":costAndCube_check,

    "['Whirpool Item']":whirlpool_check,

    "['Whirlpool Item']":whirlpool_check,

    "['Capacity Check']":highCube_check,

    "['Appliance Check']":appliance_check,

#    "['Charlotte Pipe Hold']",

#    "['Possible One-Time Buy']",

#    "['Decrease Stock Level Check']",

    "['ICF Cost greater than $30,000']":highDollar_check,

    "['ICF Cost greater than $40,000']":highDollar_check,

    # "['Capacity Check & ICF Cost greater than 40000']",

    # "['Duplicate Request - Original ICF# 1635 executed']",

    # "['Changes to this USN/DC need review due to an ongoing CM initiative']",

    # "['Duplicated demand on ICF. ICF processed manually with correct demand']"}

    }

 

    ## from sheet processing script

gbq_export = pd.read_csv(r'<LOC_ON_DRIVE>')

 

ICFs = list(gbq_export.ICF_Number.unique())

saveLoc = "<LOC_ON_DRIVE>"

seenFile = open(r'<LOC_ON_DRIVE>', "a+")

seenContent = open(r'<LOC_ON_DRIVE>', "r").read()

seenList = seenContent.split()

 

# put into email format

for ICF in ICFs:

    if (str(ICF) in seenList):

        continue

    else:

        df = gbq_export[gbq_export['ICF_Number'] == ICF]

        merchant = df[['Category_Merchant']]

        buyer = df[['IPR_Buyer']]

        reason=df[['Decision_Reason']]

        reason = np.array_str(df[['Decision_Reason']].iloc[0].values)

        costTotal = df['Total_ICF_Cost'].sum()

        cubeTotal = df['Total_ICF_Cube'].sum()

        me = <"EMAIL">

        seenFile.write(str(ICF))

        seenFile.write("\n")

        emailTable = df[['ICF_Number','Customer_name','Requestor', 'USN', 'Product_Description', 'Category_Level_1','Category_Level_2',               'warehouse_num','Estimated_Monthly_Units','UOM', 'Demand_Profile','Total_ICF_Cost','Total_ICF_Cube']]

        ## use decision reason to call function from dict

        reasonDict[reason]()

        emailTable.to_csv(saveLoc + str(ICF)+'.csv')

 

seenFile.close()
