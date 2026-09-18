import streamlit as st
import pandas as pd

def show_services(services) -> list:
    '''
    Displays the Ext. Services of the supplier and returns the vendor's data dict
    with the services that are to be added to the purchase order.
    '''

    services = services.to_dict()

    # Create an empty list to hold the services the user ticks
    selected_services = []

    for i in services['Ext_Service']:
        service = services['Ext_Service'][i]
        default = services['Deffault'][i]

        is_ticked = st.checkbox(label=service, value=default)

        if is_ticked:
            selected_services.append(service)

    # Return the selected services list
    return selected_services

    