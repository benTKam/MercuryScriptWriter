import re as re
from datetime import time

import pandas as pd

import time
import webbrowser

import subprocess

import webbrowser

from selenium import webdriver

from selenium.webdriver.common.keys import Keys

import openpyxl

DEFRouter = ''

DOMAinname = ''

IPMask = ''

ADDDns = ''



def filter_data(room_numbers):
    data = pd.read_excel('CCA - Erie - IP Addresses.xlsx')
    # Filter data from Cranberry sheet based on criteria
    filtered_data = data[data['Model'] == 'Mercury CCS-UC-1-X']
    filtered_data = filtered_data[filtered_data['Room Number'].isin(room_numbers)]

    # Reset the index of the filtered_data DataFrame
    filtered_data = filtered_data.reset_index(drop=True)
    return filtered_data


def getdeviceips(filtered_data):
    ipList = []
    for i, row in filtered_data.iterrows():
        ip_address = row['IP Address']
        ipList.append(ip_address)
    return ipList


def update_avf(filtered_data):
    mercury = pd.read_excel(
        r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\Mercury.xlsx",
        sheet_name='AVF')
    for i, row in filtered_data.iterrows():
        room_number = str(row['Room Number'])  # Convert room number to integer
        room_type = row['Room Type']
        hostname = row['Hostname']
        general_room_name = str(room_type) + ' ' + str(room_number)
        fusion_room_name = str(hostname)
        outlook_resource_address = 'PA16_Room_' + str(room_number) + '@' + DOMAinname
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
        room_number = str(row['Room Number'])  # Convert room number to integer
        hostname = row['Hostname']
        ip_address = row['IP Address']

        # Update the specific columns in the IPCONFIG sheet for the corresponding rows
        mercury.iloc[i, mercury.columns.get_loc('HOSTname')] = hostname
        mercury.iloc[i, mercury.columns.get_loc('IPAddr')] = ip_address
        mercury.iloc[i, mercury.columns.get_loc('DEFRouter')] = DEFRouter
        mercury.iloc[i, mercury.columns.get_loc('DOMAinname')] = DOMAinname
        mercury.iloc[i, mercury.columns.get_loc('IPMask')] = IPMask
        mercury.iloc[i, mercury.columns.get_loc('ADDDns')] = ADDDns

    # Save the updated Mercury sheet with the new data
    mercury.to_excel(
        r"C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\MercuryIP.xlsx",
        sheet_name='IPCONFIG', index=False)

def webbrowseropen(current_ips):
    for x in current_ips:
        # Open the webpage in a new tab
        webdriver_path = '/path/to/chromedriver'

        driver = webdriver.Chrome('/path/to/chromedriver')

        driver.get('http://' + x)

        time.sleep(5)
        # Find the username and password input fields and enter your credentials
        username_field = driver.find_element('cred_userid_inputtext')
        username_field.send_keys('Username')

        password_field = driver.find_element('cred_password_inputtext')
        password_field.send_keys('Passowrd')
        # Submit the login form
        password_field.send_keys(Keys.RETURN)


def getmac(current_ips, filtered_data):
    data = pd.read_excel('CCA - Erie - IP Addresses.xlsx')
    for x in current_ips:
        arp_output = subprocess.check_output(['arp', '-a', x])
        arp_output = arp_output.decode('utf-8')
        mac_address = re.search(r'([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})', arp_output)
        for i, row in data.iterrows():
            if data.iloc[i, data.columns.get_loc('IP Address')] == x:
                data.iloc[i, data.columns.get_loc('MAC Address')] = mac_address.group(0)
        data.to_excel('ErieSheet.xlsx')
        print(f"IP: {x} | MAC: {mac_address.group(0)}")


def run_ip(script_path):
    # Build the PowerShell command
    powershell_cmd = [
        'powershell.exe',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        f'Start-Process -Verb RunAs powershell -ArgumentList "-ExecutionPolicy Bypass -File {script_path}"'
    ]

    # Run the PowerShell command as a subprocess with administrative privileges
    subprocess.run(powershell_cmd, shell=True)


def run_avf(script_path):
    # Build the PowerShell command
    powershell_cmd = [
        'powershell.exe',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        f'Start-Process -Verb RunAs powershell -ArgumentList "-ExecutionPolicy Bypass -File {script_path}"'
    ]

    # Run the PowerShell command as a subprocess with administrative privileges
    subprocess.run(powershell_cmd, shell=True)


def main():
    current_ips = []
    avf_script_path = r'C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\SetupAVF.ps1'
    ip_script_path = r'C:\Users\bkamide\Downloads\Mercury_EnterpriseConfigUtility_v1.3\Mercury_EnterpriseConfigUtility_v1.3\SetupIPConfig.ps1'

    room_numbers = ['L07', 107.0, 106.0, 208.0, 'L06']  # Add your desired room numbers here
    filtered_data = filter_data(room_numbers)
    # update_avf(filtered_data)
    #update_ip(filtered_data)
    current_ips = getdeviceips(filtered_data)
    webbrowseropen(current_ips)
    # getmac(current_ips, filtered_data)
    # run_avf(avf_script_path)
    #run_ip(ip_script_path)


if __name__ == '__main__':
    main()
