import re as re

import pandas as pd

import webbrowser

import subprocess

DEFRouter = ''

DOMAinname = ''

IPMask = ''

ADDDns = ''

current_ips = []


def filter_data(room_numbers):
    data = pd.read_excel('Cranberry.xlsx')
    # Filter data from Cranberry sheet based on criteria
    filtered_data = data[data['Model'] == 'Mercury CCS-UC-1-X']
    filtered_data = filtered_data[filtered_data['Room Number'].isin(room_numbers)]

    # Reset the index of the filtered_data DataFrame
    filtered_data = filtered_data.reset_index(drop=True)
    return filtered_data


def update_avf(filtered_data):
    mercury = pd.read_excel(
        r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\Mercury.xlsx",
        sheet_name='AVF')
    for i, row in filtered_data.iterrows():
        room_number = int(row['Room Number'])  # Convert room number to integer
        room_type = row['Room Type']
        hostname = row['Hostname']
        general_room_name = str(room_type) + ' ' + str(room_number)
        fusion_room_name = str(hostname)
        outlook_resource_address = 'PA17_Room' + str(room_number) + '@' + DOMAinname
        bluetooth_friendly_name = ''

        # Determine the Bluetooth friendly name based on room type
        if room_type == 'Conference Room':
            bluetooth_friendly_name = 'Conf-' + str(room_number)
        elif room_type == 'Live Session':
            bluetooth_friendly_name = 'Live-' + str(room_number)

        # Update the specific columns in the Mercury sheet for the corresponding rows
        mercury.iloc[i, mercury.columns.get_loc('GeneralRoomName')] = general_room_name
        mercury.iloc[i, mercury.columns.get_loc('FusionRoomName')] = fusion_room_name
        mercury.iloc[i, mercury.columns.get_loc('OutlookResourceAddress')] = outlook_resource_address
        mercury.iloc[i, mercury.columns.get_loc('BluetoothFriendlyName')] = bluetooth_friendly_name

    # Save the updated Mercury sheet with the new data
    mercury.to_excel(
        r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\Mercury.xlsx",
        sheet_name='AVF', index=False)


def update_ip(filtered_data):
    mercury = pd.read_excel(
        r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\MercuryIP.xlsx",
        sheet_name='IPCONFIG')
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

        current_ips.append(ip_address)




    # Save the updated Mercury sheet with the new data
    mercury.to_excel(
        r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\MercuryIP.xlsx",
        sheet_name='IPCONFIG', index=False)


def webbrowseropen():
    for x in current_ips:
        webbrowser.open_new_tab('http://' + x)

def getmac():
    for x in current_ips:
        arp_output = subprocess.check_output(['arp', '-a', x])
        arp_output = arp_output.decode('utf-8')
        mac_address = re.search(r'([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})', arp_output)
        print(f"IP: {x} | MAC: {mac_address.group(0)}")
def main():
    room_numbers = [133.0, 225.0]  # Add your desired room numbers here
    filtered_data = filter_data(room_numbers)
    update_avf(filtered_data)
    update_ip(filtered_data)
    #webbrowseropen()
    getmac()


if __name__ == '__main__':
    main()
