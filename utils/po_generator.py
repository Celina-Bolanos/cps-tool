from pathlib import Path
from openpyxl import load_workbook
import io
import pandas as pd
from datetime import datetime
from copy import copy
from openpyxl.utils import range_boundaries
from openpyxl.styles import Alignment

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_PATH = SCRIPT_DIR / 'po_templates' / 'po_template.xlsx'

MASTER_MAPPING_PO = {
    "cvn_number_cell": "B4",
    "model_name_cell": "J5", # Model #
    "model_code_cell": "J6",
    "total_vehicles_cell": "P6",    # Total cars on the job
    "accessory_start_row": 10,
    "accessory_code_col" : 2, # Column B
    "accessory_qty_col": 3, # Column C
    "accessory_desc_col" :  5, # Column E
    "accessory_fitting_time_col" :  12, # Column L
}

# Function to select supplier data from database
def vendor_data(subcontractor_name: str) -> dict:
    ''' Reads subcontractors data and extract the data for the selected supplier
    Args
        subcontractor_name - str: name of the selected subcontractor

    Returns
        vendor_data - dict: all de details for the selected supplier

    '''

    subcontractor_data = pd.read_excel(SCRIPT_DIR.parent / 'data' / 'subcontractors.xlsx', sheet_name='Subcontractors')
    subcontractor_data = subcontractor_data[subcontractor_data['name'] == subcontractor_name]

    subcontractor_prices = pd.read_excel(SCRIPT_DIR.parent / 'data' / 'subcontractors.xlsx', sheet_name='Ext_services')
    subcontractor_prices = subcontractor_prices[subcontractor_prices['name'] == subcontractor_name]


    return subcontractor_data, subcontractor_prices


# Function to copy/paste the formating of a row
def copy_row_format(ws, source_row: int, target_row: int):
    """Copies all cell formatting from a source row to a target row.
    
    Args
    ws - excel worksheet to copy fromat from
    source_row - int: row number to copy from
    target_row - int: row number to apply format to

    """
    # Loop over every active column in the spreadsheet
    for col in range(1, ws.max_column + 1):
        source_cell = ws.cell(row=source_row, column=col)
        target_cell = ws.cell(row=target_row, column=col)
        
        # Explicitly copy styles over to the new cell if they exist
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.alignment = copy(source_cell.alignment)


# Function to add new rows  
#def adjust_rows(ws, num_rows: int, default_rows: int = 4, base_row: int = 23):

def adjust_rows(ws, rows_to_add: int, base_row: int = 23):
    """Inserts extra rows safely, clones formatting from a base template row, 
    and handles complex cell merges without destroying neighboring rows.

    Args
        ws - excel worksheet: ws to be modified
        rows_to_add -int: number of rows to be inserted
        base_row -int: row number after which rows are to be inserted
    """
    if rows_to_add <= 0:
        return

    template_height = ws.row_dimensions[base_row].height
    max_col = ws.max_column

    # --- STEP 1: ISOLATE & CLEANLY DETACH LOWER MERGES ---
    # Tracking arrays to avoid mutating ws.merged_cells during iteration
    row_merges_to_clone = []
    lower_merges_to_shift = []
    
    # We create a static snapshot list of current merge strings
    current_merge_strings = [str(r) for r in ws.merged_cells.ranges]

    for merge_str in current_merge_strings:
        min_col, min_row, max_col_bound, max_row = range_boundaries(merge_str)
        
        # Capture the merge format pattern from the base row (e.g., C to H)
        if min_row == base_row and max_row == base_row:
            row_merges_to_clone.append((min_col, max_col_bound))
            
        # Target only merges strictly below our insertion threshold
        elif min_row > base_row:
            lower_merges_to_shift.append((min_col, min_row, max_col_bound, max_row))
            # Completely clean out the old merge reference before inserting rows
            ws.unmerge_cells(merge_str)

    # --- STEP 2: SAFE ROW INSERTION ---
    # Now that lower merges are detached, Excel won't stretch or warp them
    ws.insert_rows(base_row + 1, amount=rows_to_add)

    # --- STEP 3: APPLY FORMATTING & MERGES TO NEWLY CREATED ROWS ---
    for i in range(1, rows_to_add + 1):
        target_row = base_row + i
        
        # Ensure row height is explicitly set on EVERY new row, including the final ones
        ws.row_dimensions[target_row].height = template_height
        
        # Clone cell-by-cell layout and borders across the entire grid width
        for col_idx in range(1, max_col + 1):
            source_cell = ws.cell(row=base_row, column=col_idx)
            target_cell = ws.cell(row=target_row, column=col_idx)
            
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.border = copy(source_cell.border)
                target_cell.fill = copy(source_cell.fill)
                target_cell.number_format = copy(source_cell.number_format)
                target_cell.alignment = copy(source_cell.alignment)

        # Apply the cloned horizontal merges (e.g., C to H) to this specific new row
        for start_col, end_col in row_merges_to_clone:
            ws.merge_cells(
                start_row=target_row, 
                start_column=start_col, 
                end_row=target_row, 
                end_column=end_col
            )
            
            # CRITICAL FIX FOR BORDERS IN MERGED CELLS: 
            # Openpyxl requires border styles applied to the outer edges of the merge track
            for c_idx in range(start_col + 1, end_col + 1):
                edge_cell = ws.cell(row=target_row, column=c_idx)
                edge_cell.border = copy(ws.cell(row=target_row, column=start_col).border)

    # --- STEP 4: RESTORE SHIFTED LOWER MERGES ---
    # Re-link the original lower blocks at their newly shifted positions
    for min_col, min_row, max_col_bound, max_row in lower_merges_to_shift:
        ws.merge_cells(
            start_row=min_row + rows_to_add,
            start_column=min_col,
            end_row=max_row + rows_to_add,
            end_column=max_col_bound
        )

            
def find_vins_row(ws, search_text: str = 'VIN-Nrs'):
    """Loops through Column A to find a cell containing specific text.
    
    Returns the integer row number if found, or None if not found.
    """
    search_text_lower = search_text.strip().lower()

    # Search through Column A (Column 1)
    for row_idx in range(1, ws.max_row + 1):
        cell_value = ws.cell(row=row_idx, column=1).value
        
        if cell_value is not None:
            if search_text_lower in str(cell_value).strip().lower():
                return row_idx  # Return the row number immediately
                
    return None

# Function that reads the worsheet and returns the useful information
def collect_data(uploaded_ws, mapping: dict) -> pd.DataFrame:
    """
    Reads data from the worksheet, extracts relevant information, 
    and returns the collected relevant information.
    
    Args:
        uploaded_ws str: string path to the uploaded worksheet file.

    Returns:
        collected_data -dict: basic data from the worksheet
        accessories_dict -dict: code + description of all accessories
        num_rows -int: number of accessories/rows to be added to the PO template
        vins_ors_dict -dict: dictionary with collected VIN numbers as keys and their corresponding OR number as values
    """
    # 1. Open the uploaded master worksheet
    # data_only=True ensures we read final values, not active Excel formulas
    workbook = load_workbook(uploaded_ws, data_only=True)
    worksheet = workbook["Worksheet"] #Target the sheet that contains the data
   
    # 2. Extract the header details from the top cells
    collected_data = {
        'cvn_num' : worksheet[mapping['cvn_number_cell']].value or 'NO_REF_FOUND',
        'vehicles_qty' : worksheet[mapping["total_vehicles_cell"]].value or 0,
        'model_name' : worksheet[mapping['model_name_cell']].value or 'NO_MODEL_FOUND',
        'model_code' : worksheet[mapping['model_code_cell']].value or 'NO_MODEL_CODE_FOUND'
        # Add mapping to CPS Stickers
    }

    # 3. Read worksheet, find and count number of accessories and description
    current_row = MASTER_MAPPING_PO["accessory_start_row"] #Start checking on row 10
    accessories_dict = {}  # Dictionary to hold accessory codes and their descriptions
    blank_rows = 0

    while True:
        # 1. Read the value from B2 and E10
        code_value = worksheet.cell(row=current_row, column=mapping["accessory_code_col"]).value #mapping cell B10
        accessory_desc = worksheet.cell(row=current_row, column=mapping["accessory_desc_col"]).value #mapping cell E10
        qty_vin = worksheet.cell(row=current_row, column=mapping["accessory_qty_col"]).value #mapping cell C10
        fitting_time = worksheet.cell(row=current_row, column=mapping["accessory_fitting_time_col"]).value or 0 #mapping cell L10

        # 2. Stop if two consecutive rows are blank
        if blank_rows >= 2:  
                break
        
        # 3. Handle blank cells  
        if code_value is None or str(code_value).strip() == '':
            blank_rows += 1
            current_row += 1
            continue  

        # 5. Process valid rows
        else:
            blank_rows = 0  # Reset blank row counter
            accessories_dict[code_value] = {
                "description": accessory_desc,
                "quantity": qty_vin,
                "fitting_time": fitting_time
            }
            current_row += 1
   
        num_rows = len(accessories_dict)

    # 4. Find and save the VIN numbers + OR numbers

    # First, find VINs cell location:
    vins_row = find_vins_row(worksheet)+ 1
    vins_ors_dict = {}
    blank_rows_vins = 0

    # Read rows below to find all vins and or numbers and add them to the dict
    while True:
            # 1. Read the values for the VINs and OR numbers
            vin_num = worksheet.cell(row=vins_row, column=1).value #mapping first cell that contains the vins
            OR_num = worksheet.cell(row=vins_row, column=4).value #mapping cell that contains the OR
    
            # 2. Stop if two consecutive rows are blank
            if blank_rows_vins >= 2:  
                    break
            
            # 3. Handle blank cells  
            if vin_num is None or str(vin_num).strip() == '':
                blank_rows_vins += 1
                vins_row += 1
                continue  
    
            # 4. Process valid rows
            else:
                blank_rows_vins = 0  # Reset blank row counter
                vins_ors_dict[vin_num] = OR_num # add VIN + or to dictionary
                vins_row += 1 # Go to next row
    

    return collected_data, accessories_dict, num_rows, vins_ors_dict


def fill_template(TEMPLATE_PATH: str, supplier: str, collected_data: dict, accessories_dict: dict, vins_ors_dict):
    '''
    Fills the template based on the supplier and returns the filled document.
    
    Args:
        template_path -str: string path to po template
        supplier -str: The selected supplier.
        collected_data (dict): Dictionary containing header details.
        accessories_dict (dict): Dictionary containing accessory details.
        num_rows (int): Number of accessory rows.
    Returns:
        filled_template: The filled template document.
    '''
    # 1. Load the template
    po_temp_wb = load_workbook(TEMPLATE_PATH, data_only=True)
    po_template = po_temp_wb.active

    subcontractor_data, subcontractor_prices = vendor_data(supplier)
    
    
    # 2. Fill in the header details
    po_template['B6'] = subcontractor_data['name_long'].iloc[0]
    po_template['B7'] = f"{subcontractor_data['address_1'].iloc[0]}\n{subcontractor_data['address_2'].iloc[0]}"
    po_template['L4'] = datetime.now().strftime('%d/%m/%Y')
    po_template['L8'] = collected_data['cvn_num']
    po_template['B13'] = f"{collected_data['vehicles_qty']}x TOY {collected_data['model_name']} \n {collected_data['model_code']}"
    po_template['K19'] = subcontractor_data['hourly_rate'].iloc[0]

    # 3. Determine and add new rows if needed for the OR/VIN numbers
    total_vins = collected_data['vehicles_qty']
    
    if total_vins >= 10:
        # Calculate how many rows to add for the VINs and add them
        extra_rows_vins = (total_vins - 10) // 3 + 1
        adjust_rows(po_template, rows_to_add=extra_rows_vins, base_row=15)

        # Merge the cells on column B, where qty, model and model code are mentioned
        po_template.merge_cells(start_row=13, start_column=2, end_row=15 + extra_rows_vins, end_column=2)
        merged_cell_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Center those cells starting on the base row 13 (see template B13)
        for r_idx in range(13, 15 + extra_rows_vins):
            po_template.cell(row=r_idx, column=2).alignment = merged_cell_alignment

    else:
        extra_rows_vins = 0 #  No rows added for vin/or numbers
        # Merge and center 'Make & model' area (B13 to B15)
        po_template.merge_cells(start_row=13, start_column=2, end_row=15, end_column=2)
        merged_cell_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for r_idx in range(13, 15):
            po_template.cell(row=r_idx, column=2).alignment = merged_cell_alignment

        
    # 4. Fill in VIN numbers
    vins_row_start = 13
    max_rows_per_col = 3 + extra_rows_vins  # Total physical rows available per column block

    # Map out exactly which numerical columns openpyxl needs to write to:
    # Column map syntax: (VIN_Start_Col, VIN_End_Col, OR_Col)
    column_layout_groups = [
        (3, 5, 6),   # Group 1: VINs in C-E (3-5), OR in F (6)
        (7, 9, 10),  # Group 2: VINs in G-I (7-9), OR in J (10)
        (11, 13, 14) # Group 3: VINs in K-M (11-13), OR in N (14)
    ]

    # 5. POPULATE DATA VERTICALLY THEN HORIZONTALLY
    # Loop through the dictionary items (vin, or) extracted previously
    for idx, (vin_num, or_num) in enumerate(vins_ors_dict.items()):

        # Calculate row offset downward within the current column capacity
        row_offset = idx % max_rows_per_col
        target_row = vins_row_start + row_offset

        # Calculate which of the 3 column groups we are currently dropping data into
        group_idx = idx // max_rows_per_col

        # If the file has more items than the grid structure can physically hold, prevent an index crash
        if group_idx >= len(column_layout_groups):
            print(f"Warning: Maximum template space reached. Skipping VIN {vin_num}")
            break

        # Extract the target column indices for this group
        vin_start_col, vin_end_col, or_col = column_layout_groups[group_idx]

        # Write the VIN number to the left-most cell of the merge block
        po_template.cell(row=target_row, column=vin_start_col, value=vin_num)

        # Re-merge the C-E, G-I, or K-M cells for this newly populated target row
        po_template.merge_cells(
            start_row=target_row, start_column=vin_start_col, 
            end_row=target_row, end_column=vin_end_col
        )

        # Write the corresponding OR number directly into the adjacent single column (F, J, or N)
        po_template.cell(row=target_row, column=or_col, value=or_num)



    # 5. Determine and add new rows if needed for the items
    total_items = len(accessories_dict)
    if  total_items> 4:
        extra_rows_items = total_items - 4
        adjust_rows(po_template, rows_to_add=extra_rows_items, base_row=23 + extra_rows_vins)
    else:
        extra_rows_items = 0

    
    # Now fill accessories list as of row 20 + extra added rows in the VINs area
    start_row = 20 + extra_rows_vins

    # Define variable to collect intallation costs per item
    total_per_item_list = []
    
    for idx, (key, item_data) in enumerate(accessories_dict.items()):
        current_target_row = start_row + idx
    # item_data = dict of each accessory. Eg 3585 : {'description': 'Pintle hook'}

        # Define variable to be able to refer to them later.
        # At least supplier fitting_time needs a variable name to then be added
        acc_code = key # Column B
        acc_desc = item_data.get('description', 'Not found')    # Column C
        qty_vin =  item_data.get('quantity', 0)
        fitting_time = item_data.get('fitting_time')  # Column J
        supplier_price = subcontractor_data['hourly_rate'].iloc[0]
        price_vin = fitting_time * supplier_price
        total = fitting_time * supplier_price * collected_data['vehicles_qty']
        total_per_item_list.append(total) # Add price to list of instalattion costs
        
        # Write to specific columns based on your template's layout
        # (Remember: use the top-left cell coordinate if the column is merged!)
        po_template.cell(row=current_target_row, column=2).value = acc_code
        po_template.cell(row=current_target_row, column=3).value = acc_desc
        po_template.cell(row=current_target_row, column=9).value = qty_vin     # Column I
        po_template.cell(row=current_target_row, column=10).value = fitting_time
        po_template.cell(row=current_target_row, column=11).value = price_vin
        po_template.cell(row=current_target_row, column=12).value = total 

        # Fix unit price if fitting time == '-'
        if type(price_vin) == str:
            price_vin = '-'
        else:
            price_vin = price_vin
        # Fix total price if fitting time == '-'
        if type(total) == str:
            total  = '-'
        else:
            total = total

    # Calculate row number for PO subtotal with base row 27
    subtotal_row = 27 + extra_rows_vins + extra_rows_items
    subtotal_cell = f'L{subtotal_row}'

    # Add fill in subtotal cell with total installation costs
    subtotal = sum(total_per_item_list)
    po_template[subtotal_cell] = subtotal

    # Calculate VAT
    vat_row = subtotal_row + 3 # first row below total
    vat_cell = f'L{vat_row}'
    vat = 0 # subtotal  * 0.21 #vat yet to be calculated
    #po_template[vat_cell] = 

    # Calculate PO grand total
    grand_total_row = vat_row + 3 # row below vat
    grand_total_cell = f'L{grand_total_row}'
    po_template[grand_total_cell] = subtotal + vat

    # Convert to 
    po_stream = io.BytesIO()
    po_temp_wb.save(po_stream)
    po_stream.seek(0)

    return po_stream


####### NOTES
# Maybe it would be a good idea to make a function that just scans the worksheet to determine
# if new rows need to be added, copy rows format, adds new rows and returns a modified template
# ready to be filled in.