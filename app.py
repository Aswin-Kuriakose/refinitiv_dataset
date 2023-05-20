import json
import requests
import string
import time
from clients import getGSpreadClient
from create_worksheet import create_sheet


# letter_list contains the alphabets. It is used to loop over the column of the google sheet


letter = string.ascii_uppercase
letter_list = list(letter)

# getGSpreadClient contains the credential for using the gspread library to modify the sheet

client = getGSpreadClient()
spread_sheet = client.open("REFINITIV DATA")

refinitive_companies_url = "https://www.refinitiv.com/bin/esg/esgsearchsuggestions"
resoponse = requests.get(refinitive_companies_url)
data = json.loads(resoponse.text)

sheet_no = len(spread_sheet.worksheets())

if sheet_no > 0:
    sheet = create_sheet()

# '''
# ========================================================================================================================

for i, item in enumerate(data):

    time.sleep(1)

    ricCode = item["ricCode"]
    print(i + 1)

    params = {"ricCode": ricCode}

    try:
        r = requests.get(url="https://www.refinitiv.com/bin/esg/esgsearchresult", params=params)
        company = r.json()

        company_name = item['companyName']
        esgScore = company["esgScore"]["TR.TRESG"]["score"]
        eScore = company["esgScore"]["TR.EnvironmentPillar"]["score"]
        sScore = company["esgScore"]["TR.SocialPillar"]["score"]
        gScore = company["esgScore"]["TR.GovernancePillar"]["score"]

        data = [[company_name,
                 ricCode,
                 esgScore,
                  eScore,
                  sScore,
                  gScore]]

        no_of_data = len(data[0])
        last_element = letter_list[no_of_data - 1]

        cell_range = sheet.range(f'A{i+2}:{last_element}{i+2}')
        
        for i,cell in enumerate(cell_range):
          
           cell_range[i].value = data[0][i]
      
        sheet.update_cells(cell_range)
      
    except:
       print( " refintiv api not working ")


# '''