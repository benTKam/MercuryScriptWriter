import re as re

import pandas as pd

import webbrowser

import subprocess

rooms = [133.0, 225.0]

data = pd.read_excel('Cranberry.xlsx')
mercury = pd.read_excel(r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\MercuryIP.xlsx", sheet_name='IPCONFIG')

# Filter data from Cranberry sheet based on criteria
filtered_data = data[data['Model'] == 'Mercury CCS-UC-1-X']
filtered_data = filtered_data[filtered_data['Room Number'].isin(rooms)]

# Reset the index of the filtered_data DataFrame
filtered_data = filtered_data.reset_index(drop=True)

# Iterate over the filtered data and update the specific columns in the IPCONFIG sheet
for i, row in filtered_data.iterrows():
    room_number = int(row['Room Number'])  # Convert room number to integer
    hostname = row['Hostname']
    ip_address = row['IP Address']

    # Update the specific columns in the IPCONFIG sheet for the corresponding rows
    mercury.iloc[i, mercury.columns.get_loc('HOSTname')] = hostname
    mercury.iloc[i, mercury.columns.get_loc('IPAddr')] = ip_address
    mercury.iloc[i, mercury.columns.get_loc('DEFRouter')] = DEFRouter
    mercury.iloc[i, mercury.columns.get_loc('DOMAinname')] = DOMAinname
    mercury.iloc[i, mercury.columns.get_loc('IPMask')] = IPMask
    mercury.iloc[i, mercury.columns.get_loc('ADDDns')] = ADDDns

    #webbrowser.open_new_tab('http://' + ip_address)

    #arp_output = subprocess.check_output(['arp', '-a', ip_address])
    #arp_output = arp_output.decode('utf-8')
    #mac_address = re.search(r'([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})', arp_output)

    #print(f"IP: {ip_address} | MAC: {mac_address.group(0)}")

# Save the updated Mercury sheet with the new data
mercury.to_excel(r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\MercuryIP.xlsx", sheet_name='IPCONFIG', index=False)

