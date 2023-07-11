import pandas as pd

rooms = [231.0, 235.0, 112.0, 208.0]

data = pd.read_excel('Cranberry.xlsx')
mercury = pd.read_excel('Mercury.xlsx', sheet_name='AVF')

# Filter data from Cranberry sheet based on criteria
filtered_data = data[data['Model'] == 'Mercury CCS-UC-1-X']
filtered_data = filtered_data[filtered_data['Room Number'].isin(rooms)]

# Reset the index of the filtered_data DataFrame
filtered_data = filtered_data.reset_index(drop=True)

# Iterate over the filtered data and update the specific columns in the Mercury sheet
for i, row in filtered_data.iterrows():
    room_number = int(row['Room Number'])  # Convert room number to integer
    room_type = row['Room Type']
    hostname = row['Hostname']
    general_room_name = str(room_type) + ' ' + str(room_number)
    fusion_room_name = str(hostname)
    outlook_resource_address = 'PA17_Room' + str(room_number) + '@ccaeducate.me'
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
mercury.to_excel('UpdatedMercury.xlsx', index=False)
