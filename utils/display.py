import streamlit as st

def show_services(services: dict, vendor_data: dict):
    '''
    Displays the Ext. Services of the supplier and returns the vendor's data dict
    with the services that are to be added to the purchase order.
    '''

    #Create a sub-dictionary or list inside vendor_data to hold the chosen services
    vendor_data['selected_services'] = []

    for key, val in services.items():
        # Set the starting state based on your services dict
        default_value = False if val == 'no' else True
        
        # Capture the live user interaction (True/False)
        is_ticked = st.checkbox(label=key, value=default_value)
        
        # If the user checked it, save it to the vendor_data
        if is_ticked:
            vendor_data['selected_services'].append(key)

    # Return the updated dictionary
    return vendor_data

    