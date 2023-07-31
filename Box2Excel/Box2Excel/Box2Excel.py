import os
import datetime
from boxsdk import Client, OAuth2
import pandas as pd

# Box API credentials
CLIENT_ID = 'YOUR_CLIENT_ID'
CLIENT_SECRET = 'YOUR_CLIENT_SECRET'
ACCESS_TOKEN = 'YOUR_ACCESS_TOKEN'

# Create an authenticated Box client
auth = OAuth2(client_id=CLIENT_ID, client_secret=CLIENT_SECRET, access_token=ACCESS_TOKEN)
box_client = Client(auth)

# Box folder ID from which you want to pull information
BOX_FOLDER_ID = 'YOUR_BOX_FOLDER_ID'

# Function to pull data from Box and save it to Excel
def pull_data_and_save_to_excel():
    # Get today's date
    today_date = datetime.datetime.now().strftime('%Y-%m-%d')

    # Pull data from Box folder
    box_folder = box_client.folder(folder_id=BOX_FOLDER_ID).get()
    items = box_folder.get_items()

    # Create a Pandas DataFrame with the information (change this according to your data structure)
    data = []
    for item in items:
        data.append([item.name, item.size, item.created_at, item.modified_at])
    
    df = pd.DataFrame(data, columns=['Name', 'Size', 'Created At', 'Modified At'])

    # Create an Excel writer with openpyxl engine
    excel_file_path = f'GE_LTDT_{today_date}.xlsx'
    writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')

    # Write DataFrame to a new tab in the Excel file
    df.to_excel(writer, sheet_name='Data', index=False)

    # Save and close the Excel file
    writer.save()
    writer.close()

# Run the function to pull data and save to Excel
pull_data_and_save_to_excel()

# Check if the month has changed and create a new file
current_month = datetime.datetime.now().strftime('%Y-%m')
if not os.path.exists(f'GE_LTDT_{current_month}.xlsx'):
    os.rename(f'GE_LTDT_{today_date}.xlsx', f'GE_LTDT_{current_month}.xlsx')

