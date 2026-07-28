import streamlit as st
from openpyxl import load_workbook
import io
import os
from utils.scanlist_maker import (
    read_ws, 
    MASTER_MAPPING, 
    xlookup_accessories, 
    xlookup_items
)
from utils.po_generator import(
    MASTER_MAPPING_PO,
    TEMPLATE_PATH,
    collect_data, 
    fill_template
)

# --- STREAMLIT WEB INTERFACE ---
st.set_page_config(page_title='Worksheet Processing Tool', page_icon='📝')
st.title('📝 Worksheet Processing Tool')
st.write('Testing the automation of scan list and PO creation.')


# 1. File Upload Box
uploaded_file = st.file_uploader('Upload your Worksheet (.xlsx)', type=['xlsx'])

# 2. Initialize session state for showing the PO form
if 'show_po_form' not in st.session_state:
    st.session_state.show_po_form = False

# 3. Process Execution
if uploaded_file is not None:
    st.success('Master worksheet loaded successfully!')
    st.write('What would you like to do?')

# 4. Make Scan List     
    if st.button('🚀 Make Scan list', use_container_width=True):
        try:
            # Run worksheet reading function
            scan_list_df = read_ws(uploaded_file, MASTER_MAPPING)
            scan_list = xlookup_accessories(scan_list_df)
            final_scan_list, ref_num = xlookup_items(scan_list) 
           
            # Show a brief preview of what we captured on screen so you can verify it
            st.info('Scan list has been generated successfully!')
           
            # Offer the file back to the browser for instant download
            st.download_button(
                label='📥 Download scan list',
                # 1. Added sep=',' and prepend the UTF-8 BOM (\ufeff)
                data=('\ufeff' + final_scan_list.to_csv(index=False, header=False, sep=';')).encode('utf-8'),
                file_name=f'scan list {ref_num}.csv',
                # 2. Fixed the MIME type specifically for CSV data
                mime='text/csv',
                use_container_width=True
            )
        
        except Exception as e:
            st.error(f'An unexpected error occurred: {e}')

# 5. Make Purchase Order        
    if st.button('🚀 Make Purchase Order', use_container_width=True):
        st.session_state.show_po_form = True

    # Keep UI rendering alive across supplier changes or button downloads
    if st.session_state.show_po_form:
        supplier = st.selectbox(
            'Select supplier', 
            options=['Select supplier', 'Operadora Pajarito', 'Taller Don Jose', 'AET', 'Carcenter', 'ICO', 'MSE']
        )
        
        if supplier != 'Select supplier':
            st.write(f'You selected: **{supplier}**')
            st.write('Generating Purchase Order...')

            try:
                collected_data, accessories_dict, num_rows = collect_data(uploaded_file, MASTER_MAPPING_PO)
                ref_num = collected_data.get('cvn_num', 'UNKNOWN')
                
                new_po_stream = fill_template(TEMPLATE_PATH, supplier, collected_data, accessories_dict, num_rows)
                #new_po_stream = fill_template(TEMPLATE_PATH, collected_data, accessories_dict, num_rows)

                # Offer the file back to the browser for instant download
                st.download_button(
                label='📥 Download PO',
                data=new_po_stream,
                file_name=f'PO_{ref_num}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True
            )

            except Exception as e:
                st.error(f'An unexpected error occurred: {e}')
                print(e)
        
                    

