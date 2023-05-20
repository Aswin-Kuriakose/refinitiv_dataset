from clients import getGSpreadClient
from datetime import date 

def create_sheet():

    today = date.today()
    new_sheet_name = today.strftime("%d-%Y-%m")

    client = getGSpreadClient()

    spread_sheet = client.open("REFINITIV DATA")
    new_sheet =  spread_sheet.add_worksheet(new_sheet_name, 12000, 6)

    default_header = [
                    'COMPANY NAME',
                    'RIC CODE',
                    'Refinitiv ESG Rating',
                    'Refinitiv E Score',
                    'Refinitiv S Score',
                    'Refinitiv G Score' 
                    ]
    
    new_sheet.insert_row(default_header,1)
    return spread_sheet.worksheet(new_sheet_name)
    
