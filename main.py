import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment # Import specific style components
from openpyxl.styles import Color # Import Color for defining colors
from openpyxl.comments import Comment # Import Comment for cell comments
import os
import re # Import the regular expression module
import itertools # Import itertools for grouping
import sys # Import sys module to detect if running as bundled executable

# Global variable for worklist file path
worklist_file_path = None
# Global variable for lookup data (from Respone - Do not Delete.xlsx)
response_lookup_data = {
    'AB_lookup': {}, # For J to K (Col A: Col B in Response file)
    'CD_lookup': {}  # For S to L (Col C: Col D in Response file)
}

# Global variable for the selected sheet name
selected_sheet_name = None # This will be a tk.StringVar

# Global variable for the sheet selection UI elements
sheet_selection_frame = None
sheet_option_menu = None

# Global variable for the highlighting option state
enable_highlight_var = None # This will be a tk.BooleanVar
# Global variable for the option to include (not skip) rows with strikethrough
include_strikethrough_rows_var = None # This will be a tk.BooleanVar

# Global variable to store template rows data (including styles and merged cells)
template_rows_data = []

# Global variable for the status label
status_label = None

def get_resource_path(relative_path):
    """
    Get the absolute path to resource, works for dev and for PyInstaller.
    """
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS.
        base_path = os.path.dirname(sys.executable)
    else:
        # If run using Python interpreter
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def select_excel_file():
    """
    Opens a file dialog for the user to select an Excel file (Worklist File)
    and displays the selected file path. Then, it populates a dropdown
    with sheet names for the user to choose from.
    """
    global worklist_file_path, selected_sheet_name, sheet_selection_frame, sheet_option_menu

    new_worklist_file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ Worklist Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if new_worklist_file_path:
        worklist_file_path = new_worklist_file_path # Update global path

        # Display only the base name of the selected file
        file_path_label.config(text=f"ไฟล์ Worklist: {os.path.basename(worklist_file_path)}")

        try:
            # Load the workbook to get sheet names.
            # We need to load in read-write mode to check for strikethrough later.
            temp_workbook = openpyxl.load_workbook(worklist_file_path) 
            sheet_names = temp_workbook.sheetnames
            
            # Attempt to get the active sheet's title. If it fails, default to the first sheet.
            active_sheet_title = None
            try:
                active_sheet_title = temp_workbook.active.title
            except Exception:
                pass # Ignore error if cannot determine active sheet

            temp_workbook.close() # Close the workbook immediately

            if not sheet_names:
                messagebox.showwarning("คำเตือน", "ไฟล์ Excel ที่เลือกไม่มีชีท!")
                file_path_label.config(text="ยังไม่ได้เลือกไฟล์ Worklist")
                convert_button.config(state=tk.DISABLED)
                # Destroy existing sheet selection elements if any
                if sheet_selection_frame:
                    sheet_selection_frame.destroy()
                    sheet_selection_frame = None
                return

            # If sheet selection frame already exists, destroy it to recreate
            if sheet_selection_frame:
                sheet_selection_frame.destroy()

            # Create or recreate the frame for sheet selection
            sheet_selection_frame = tk.LabelFrame(root, text="2. เลือกชีท", padx=10, pady=10)
            sheet_selection_frame.pack(pady=5, padx=20, fill="x")

            sheet_label = tk.Label(sheet_selection_frame, text="เลือกชีท:")
            sheet_label.pack(side=tk.LEFT, padx=(0, 10))

            selected_sheet_name = tk.StringVar(root)
            # Set initial value to the active sheet title or the first sheet name
            if active_sheet_title and active_sheet_title in sheet_names:
                selected_sheet_name.set(active_sheet_title)
            else:
                selected_sheet_name.set(sheet_names[0]) # Fallback to first sheet name

            sheet_option_menu = tk.OptionMenu(sheet_selection_frame, selected_sheet_name, *sheet_names)
            sheet_option_menu.pack(side=tk.LEFT, fill="x", expand=True)
            
            convert_button.config(state=tk.NORMAL) # Enable Convert button after file and sheet are ready

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถอ่านไฟล์ Excel ได้: {e}\n"
                                               f"โปรดตรวจสอบว่าไฟล์ไม่ได้ถูกเปิดอยู่หรือเสียหาย")
            file_path_label.config(text="ยังไม่ได้เลือกไฟล์ Worklist")
            convert_button.config(state=tk.DISABLED)
            if sheet_selection_frame:
                sheet_selection_frame.destroy()
                sheet_selection_frame = None
            worklist_file_path = None # Reset path on error

    else: # User cancelled file selection
        file_path_label.config(text="ยังไม่ได้เลือกไฟล์ Worklist")
        convert_button.config(state=tk.DISABLED)
        if sheet_selection_frame:
            sheet_selection_frame.destroy()
            sheet_selection_frame = None
        worklist_file_path = None # Reset path


def split_j_column_data(text_data):
    """
    Splits text data from Column J based on '.-', '. ', ',', or '/' delimiters.
    Ensures each split part ends with a dot if it represents a meaningful segment.
    Handles leading/trailing pipe characters and re-adds them to split parts.
    """
    if not isinstance(text_data, str):
        return [text_data] # If not a string, return as is in a list (e.g., None, numbers)

    original_had_pipes = text_data.startswith('|') and text_data.endswith('|')
    clean_text = text_data.strip('|') # Remove outer pipes for clean splitting

    # Define the delimiters and a unique temporary delimiter for splitting
    delimiters = [".-", ". ", ",", "/"]
    temp_delimiter = "###SPLIT_POINT###"

    processed_text = clean_text
    # Replace all specified delimiters with the unique temporary delimiter
    for delim in delimiters:
        processed_text = processed_text.replace(delim, temp_delimiter)
    
    # Split by the temporary delimiter
    # filter out any empty strings that might result from multiple delimiters next to each other
    raw_parts = [p.strip() for p in processed_text.split(temp_delimiter) if p.strip()]

    parts = []
    for part in raw_parts:
        if part: # Ensure part is not empty
            # Only add a dot if the part doesn't end with a dot, and it's not just a number (e.g., "123")
            # This logic can be adjusted if numbers should also end with a dot in specific cases.
            if not part.endswith('.'):
                parts.append(part + '.')
            else:
                parts.append(part)
    
    if not parts:
        # If no parts were found after splitting and processing, return a list with None
        # This handles cases where input was an empty string or only delimiters
        return [None] if clean_text.strip() else [None]

    # Re-add pipes if the original string had them, to each split part
    if original_had_pipes:
        parts = [f"|{p}|" for p in parts]

    return parts


def split_i_column_data(text_data):
    """
    Splits text data from Column I based on 'number dot' pattern (e.g., '1.', '2.', '10.').
    Crucially, it does NOT split if the 'number dot' pattern is found inside parentheses.
    Handles leading/trailing pipe characters.
    Also implements Process D: Removes '/.', '/', '.', or '\' if found at the end of each part.
    """
    if not isinstance(text_data, str):
        return [text_data] # If not a string, return as is in a list (e.g., None, numbers)

    original_had_pipes = text_data.startswith('|') and text_data.endswith('|')
    clean_text = text_data.strip('|') # Remove outer pipes for clean processing

    items = []
    # First, find all valid split points (N. patterns outside parentheses)
    valid_split_indices = [0] # Always start with the beginning of the string

    paren_level = 0
    for m in re.finditer(r'\d+\.', clean_text):
        match_start = m.start()
        
        # Check parenthesis level at the start of the current match
        current_paren_level = 0
        for char_idx in range(match_start):
            if clean_text[char_idx] == '(':
                current_paren_level += 1
            elif clean_text[char_idx] == ')':
                current_paren_level -= 1

        if current_paren_level == 0: # If N. is outside parentheses, it's a valid split point
            if match_start not in valid_split_indices:
                valid_split_indices.append(match_start)
    
    # Sort and remove duplicates
    valid_split_indices = sorted(list(set(valid_split_indices)))
    
    # Extract segments based on these valid split indices
    for k in range(len(valid_split_indices)):
        start_idx = valid_split_indices[k]
        end_idx = valid_split_indices[k+1] if k+1 < len(valid_split_indices) else len(clean_text)
        
        segment = clean_text[start_idx:end_idx].strip()
        if segment: # Only add non-empty segments
            items.append(segment)
    
    if not items:
        return [None]

    # --- Apply Process D: Remove '/.', '/', '.', or '\' if found at the end of each part ---
    processed_parts = []
    for part in items:
        if isinstance(part, str):
            temp_part = part.strip()

            if temp_part.endswith('/.'):
                temp_part = temp_part[:-2].strip()

            temp_part = temp_part.rstrip('/.\\') 

            processed_parts.append(temp_part)
        else:
            processed_parts.append(part)
    
    if not processed_parts:
        return [None]

    # Re-add pipes if the original string had them, to each split part
    if original_had_pipes:
        processed_parts = [f"|{p}|" for p in processed_parts]

    return processed_parts

def load_lookup_data():
    """
    Loads data from 'Respone - Do not Delete.xlsx' into a dictionary for VLOOKUP.
    Keys are from Column A and C, values are from Column B and D respectively.
    Specifically loads from the 'Respone' sheet.
    """
    global response_lookup_data
    # Initialize lookup data structure to hold two separate lookup dictionaries
    response_lookup_data = {
        'AB_lookup': {}, # For J to K (Col A: Col B in Response file)
        'CD_lookup': {}  # For S to L (Col C: Col D in Response file)
    }

    response_file_path = get_resource_path("Respone - Do not Delete.xlsx")

    if not os.path.exists(response_file_path):
        messagebox.showerror(
            "ข้อผิดพลาด",
            f"ไม่พบไฟล์ VLOOKUP: '{response_file_path}'\n"
            "กรุณาตรวจสอบว่าไฟล์ 'Respone - Do not Delete.xlsx' อยู่ในโฟลเดอร์เดียวกับโปรแกรม"
        )
        return False # Indicate failure to load

    try:
        lookup_workbook = openpyxl.load_workbook(response_file_path)
        # Explicitly select the 'Respone' sheet
        if 'Respone' in lookup_workbook.sheetnames:
            lookup_sheet = lookup_workbook['Respone']
        else:
            messagebox.showerror(
                "ข้อผิดพลาดชีท",
                f"ไม่พบชีท 'Respone' ในไฟล์ '{response_file_path}'\n"
                "โปรดตรวจสอบชื่อชีทในไฟล์ Respone - Do not Delete.xlsx"
            )
            return False

        # Iterate through rows, assuming data starts from row 1.
        # Column A is index 1, Column B is index 2. Column C is index 3, Column D is index 4.
        for row_idx in range(1, lookup_sheet.max_row + 1):
            # Load for J to K lookup (A:B)
            key_ab = lookup_sheet.cell(row=row_idx, column=1).value
            value_ab = lookup_sheet.cell(row=row_idx, column=2).value
            if key_ab is not None:
                response_lookup_data['AB_lookup'][str(key_ab).strip()] = value_ab

            # Load for S to L lookup (C:D)
            key_cd = lookup_sheet.cell(row=row_idx, column=3).value # Column C is index 3
            value_cd = lookup_sheet.cell(row=row_idx, column=4).value # Column D is index 4
            if key_cd is not None:
                response_lookup_data['CD_lookup'][str(key_cd).strip()] = value_cd
        return True # Indicate successful load
    except Exception as e:
        messagebox.showerror(
            "ข้อผิดพลาดในการโหลดไฟล์ VLOOKUP",
            f"เกิดข้อผิดพลาดขณะโหลดไฟล์ 'Respone - Do not Delete.xlsx': {e}"
        )
        return False # Indicate failure to load

def load_template_rows():
    """
    Loads the first two rows (including values, styles, comments, and merged cells)
    from the 'Template' sheet of 'Respone - Do not Delete.xlsx'.
    """
    global template_rows_data
    template_rows_data = [] # Reset template data

    response_file_path = get_resource_path("Respone - Do not Delete.xlsx")

    if not os.path.exists(response_file_path):
        messagebox.showerror(
            "ข้อผิดพลาด",
            f"ไม่พบไฟล์ Template: '{response_file_path}'\n"
            "กรุณาตรวจสอบว่าไฟล์ 'Respone - Do not Delete.xlsx' อยู่ในโฟลเดอร์เดียวกับโปรแกรม"
        )
        return False
    
    try:
        template_workbook = openpyxl.load_workbook(response_file_path)
        if 'Template' in template_workbook.sheetnames:
            template_sheet = template_workbook['Template']
        else:
            messagebox.showerror(
                "ข้อผิดพลาดชีท",
                f"ไม่พบชีท 'Template' ในไฟล์ '{response_file_path}'\n"
                "โปรดตรวจสอบชื่อชีทในไฟล์ Respone - Do not Delete.xlsx"
            )
            return False

        # Store merged cell ranges from the template sheet
        merged_cells_ranges_from_template = []
        for merged_range in template_sheet.merged_cells.ranges:
            merged_cells_ranges_from_template.append(str(merged_range)) 
        
        # Add a placeholder for merged ranges. The first item in template_rows_data
        # will now be this list of merged ranges.
        template_rows_data.append(merged_cells_ranges_from_template)

        # Read cell data for the first two rows (1-based indexing in openpyxl)
        for r_idx in range(1, 3): # Rows 1 and 2 of the template
            row_cells_data = {}
            # Iterate through columns up to the max_column of the template sheet
            for c_idx in range(1, template_sheet.max_column + 1): 
                cell = template_sheet.cell(row=r_idx, column=c_idx)
                
                # Copy cell attributes: For styles, recreate objects to avoid StyleProxy issues
                # For Comment, recreate object
                cell_data = {
                    'value': cell.value,
                    # Recreate PatternFill to avoid StyleProxy errors
                    'fill': PatternFill(start_color=cell.fill.start_color, 
                                        end_color=cell.fill.end_color, 
                                        fill_type=cell.fill.fill_type) if cell.fill else None, 
                    'font': cell.font.copy() if cell.font else None,
                    'border': cell.border.copy() if cell.border else None,
                    'alignment': cell.alignment.copy() if cell.alignment else None,
                    'number_format': cell.number_format,
                    # Recreate Comment object
                    'comment': Comment(cell.comment.text, cell.comment.author) if cell.comment else None 
                }
                row_cells_data[cell.column_letter] = cell_data # Store by column letter
            template_rows_data.append(row_cells_data)
        
        template_workbook.close() # Close the workbook
        return True
    except Exception as e:
        messagebox.showerror(
            "ข้อผิดพลาดในการโหลด Template",
            f"เกิดข้อผิดพลาดขณะโหลดชีท 'Template' จากไฟล์ 'Respone - Do not Delete.xlsx': {e}"
        )
        return False

def run_conversion_process():
    """
    Wrapper function to handle UI state before and after conversion.
    """
    global status_label

    # Disable button and show initial processing message
    convert_button.config(state=tk.DISABLED)
    status_label.config(text="กำลังแปลงข้อมูล... โปรดรอสักครู่")
    root.update_idletasks() # Update the GUI immediately

    try:
        convert_to_maximo()
    finally:
        # Re-enable button
        convert_button.config(state=tk.NORMAL)
        # Status label will be updated by convert_to_maximo in case of success/failure
        # If it was cancelled by user, it's already set there.
        root.update_idletasks()


def convert_to_maximo():
    """
    Reads data from specified columns (B, F, H, I, J) starting from Row 3 of the
    selected Worklist file and writes them to new columns (D, E, I, J, K, S) in a new Excel file.
    Column J data is split and duplicated across new rows as necessary.
    Column I data is split (excluding within parentheses and with '/.' removed from end) and duplicated across new rows.
    Process E (VLOOKUP) is applied to Column J to populate Column K, and to Column S to populate Column S (overwrite).
    Process F (Counting in G and H based on D, E, J groups) is applied.
    Process for highlighting cells in Column I (if enabled and condition met).
    Process G (Copying template rows) is applied at the beginning of the new sheet.
    The user is prompted to choose the save location for the new file.
    
    New Feature: Option to include/exclude rows with strikethrough formatting.
    """
    global selected_sheet_name # Ensure selected_sheet_name is accessible
    global enable_highlight_var # Ensure enable_highlight_var is accessible
    global include_strikethrough_rows_var # Ensure new strikethrough variable is accessible
    global status_label # Ensure status_label is accessible

    if not worklist_file_path:
        messagebox.showwarning("คำเตือน", "กรุณาเลือกไฟล์ Worklist ก่อนดำเนินการ!")
        status_label.config(text="โปรดเลือกไฟล์ Worklist")
        return
    
    if not selected_sheet_name or not selected_sheet_name.get():
        messagebox.showwarning("คำเตือน", "กรุณาเลือกชีทที่จะแปลงข้อมูล!")
        status_label.config(text="โปรดเลือกชีท")
        return

    # Try to load lookup data first
    if not load_lookup_data():
        status_label.config(text="ข้อผิดพลาดในการโหลด VLOOKUP Data")
        return # Stop if lookup data cannot be loaded

    # Try to load template rows for Process G
    if not load_template_rows():
        status_label.config(text="ข้อผิดพลาดในการโหลด Template Data")
        return # Stop if template rows cannot be loaded

    try:
        # Load the workbook in read-write mode to access formatting like strikethrough
        source_workbook = openpyxl.load_workbook(worklist_file_path)
        # Use the selected sheet name
        source_sheet = source_workbook[selected_sheet_name.get()]

        new_workbook = openpyxl.Workbook()
        new_sheet = new_workbook.active
        new_sheet.title = "Template"

        # --- Process G: Copy template rows to the new sheet ---
        # First, handle merged cells from the template
        if template_rows_data and isinstance(template_rows_data[0], list):
            for merged_range_str in template_rows_data[0]:
                try:
                    new_sheet.merge_cells(merged_range_str)
                except Exception as merge_err:
                    print(f"Warning: Could not merge cells {merged_range_str}: {merge_err}") 
        
        # Then, copy cell data for the first two rows
        # template_rows_data[0] is merged ranges, so actual row data starts from index 1.
        for r_idx_template_data in range(1, 3): # Process data for original template rows 1 and 2
            # Corresponding row in new_sheet is also r_idx_template_data (1 and 2)
            row_data_to_copy = template_rows_data[r_idx_template_data] 
            for col_letter, cell_info in row_data_to_copy.items():
                target_cell = new_sheet[f"{col_letter}{r_idx_template_data}"]
                target_cell.value = cell_info['value']
                
                # Copy styles and comment
                if cell_info['fill']:
                    target_cell.fill = cell_info['fill']
                if cell_info['font']:
                    target_cell.font = cell_info['font']
                if cell_info['border']:
                    target_cell.border = cell_info['border']
                if cell_info['alignment']:
                    target_cell.alignment = cell_info['alignment']
                if cell_info['number_format']:
                    target_cell.number_format = cell_info['number_format']
                if cell_info['comment']:
                    target_cell.comment = cell_info['comment']

        # Define a fill style for highlighting (color #ff9e00)
        highlight_fill = PatternFill(start_color='FF9E00', end_color='FF9E00', fill_type='solid')
        
        # This list will store tuples of ((D,E,J) values, row_index_in_new_sheet)
        # to facilitate grouping and counting for Process F.
        rows_for_f_process = [] 
        
        current_output_row_idx = 3 # This keeps track of the next available row in the new sheet

        for source_row_idx in range(5, source_sheet.max_row + 1):
            # Update status label with current processing row
            status_label.config(text=f"กำลังแปลงข้อมูล... (กำลังประมวลผลแถวที่ {source_row_idx})")
            root.update_idletasks() # Force GUI update

            # --- Strikethrough Check (Conditional Skip) ---
            # If 'include_strikethrough_rows_var' is False (default: skip strikethrough rows),
            # then check for strikethrough and skip the row if found.
            if not include_strikethrough_rows_var.get(): 
                skip_row_due_to_strikethrough = False
                for col_idx in range(1, source_sheet.max_column + 1): # Iterate through all columns in the row
                    cell = source_sheet.cell(row=source_row_idx, column=col_idx)
                    if cell.font and cell.font.strike:
                        skip_row_due_to_strikethrough = True
                        break # Found strikethrough, no need to check other cells in this row
                
                if skip_row_due_to_strikethrough:
                    # You can add a more detailed message here if needed, or just let it skip silently.
                    # For now, keeping the status update simple.
                    continue # Skip this row and move to the next source row

            data_col_b = source_sheet.cell(row=source_row_idx, column=2).value
            data_col_i = source_sheet.cell(row=source_row_idx, column=9).value
            data_col_f = source_sheet.cell(row=source_row_idx, column=6).value
            data_col_h_for_i = source_sheet.cell(row=source_row_idx, column=8).value
            data_col_j = source_sheet.cell(row=source_row_idx, column=10).value
            # Process A addition: Get data from Worklist Column I for output Column S (initial value)
            data_col_i_worklist = source_sheet.cell(row=source_row_idx, column=9).value # Column I is index 9 in source

            j_parts = split_j_column_data(data_col_j)
            i_parts = split_i_column_data(data_col_h_for_i)

            for j_part_item in j_parts:
                # Process E - First VLOOKUP for Column K (using J part)
                lookup_key_k = str(j_part_item).strip() if j_part_item is not None else ""
                vlookup_result_k = response_lookup_data['AB_lookup'].get(lookup_key_k, None)

                # Process E - Second VLOOKUP for Column S (using Worklist I data -> now in Col S)
                lookup_key_s = str(data_col_i_worklist).strip() if data_col_i_worklist is not None else ""
                vlookup_result_s = response_lookup_data['CD_lookup'].get(lookup_key_s, None)

                for i_part_item in i_parts:
                    # Write values to the current row in the new sheet
                    new_sheet.cell(row=current_output_row_idx, column=4).value = data_col_b # Col D
                    new_sheet.cell(row=current_output_row_idx, column=5).value = data_col_f # Col E
                    new_sheet.cell(row=current_output_row_idx, column=21).value = data_col_i # Col i type ex 6e 6m
                    
                    # Store Column I cell object to check its length for highlighting later
                    cell_i_output = new_sheet.cell(row=current_output_row_idx, column=9)
                    cell_i_output.value = i_part_item # Col I

                    new_sheet.cell(row=current_output_row_idx, column=10).value = j_part_item # Col J
                    new_sheet.cell(row=current_output_row_idx, column=11).value = vlookup_result_k # Col K
                    
                    # Write original Worklist I data to Column S (Index 19) initially
                    # If VLOOKUP result for Column S is not None, overwrite Column S with the result
                    cell_s_output = new_sheet.cell(row=current_output_row_idx, column=19)
                    if vlookup_result_s is not None:
                        cell_s_output.value = vlookup_result_s # Overwrite Col S with VLOOKUP result
                    else:
                        cell_s_output.value = data_col_i_worklist # Keep original if no lookup match
                    
                    # Apply highlighting if enabled and condition met
                    if enable_highlight_var.get(): # Check if highlighting is enabled by the user
                        # Check Column I content for length > 85
                        if isinstance(cell_i_output.value, str) and len(cell_i_output.value) > 85:
                            # Apply highlight to the cell in Column I (ONLY Column I)
                            cell_i_output.fill = highlight_fill

                    # Store (D, E, J) values and their corresponding row index in the new sheet
                    # Convert None to empty string for sorting comparison
                    rows_for_f_process.append((
                        (str(data_col_b) if data_col_b is not None else "",
                         str(data_col_i) if data_col_i is not None else "",
                         str(data_col_f) if data_col_f is not None else "",
                         str(j_part_item) if j_part_item is not None else ""), # Grouping key: (D, E, J) tuple
                        current_output_row_idx # The row index in the new sheet where this data was written
                    ))
                    
                    current_output_row_idx += 1 # Move to the next row for the new sheet

        # --- Process F: Counting in Column G and H based on D, E, J groups ---
        
        # Sort the collected data by the grouping keys (D, E, J)
        # This is crucial for itertools.groupby to correctly group consecutive identical keys.
        rows_for_f_process.sort(key=lambda x: x[0])

        # Apply the counting to the sorted and grouped data
        for group_key, group_iter in itertools.groupby(rows_for_f_process, key=lambda x: x[0]):
            current_count = 10 # Initialize count for each new group
            # group_iter yields ( (D,E,J), row_idx ) tuples for the current group
            for _, row_idx_in_new_sheet in group_iter:
                # Write the current count to Column G (index 7) and Column H (index 8)
                new_sheet.cell(row=row_idx_in_new_sheet, column=7).value = current_count
                new_sheet.cell(row=row_idx_in_new_sheet, column=8).value = current_count
                current_count += 10 # Increment by 10 for the next row in this group

        # Get the selected sheet name to use as the default file name suggestion
        default_file_name = f"{selected_sheet_name.get()} converted.xlsx"

        # Open a "Save As" dialog for the user to choose the save location for the new file
        output_file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", # Default file extension
            filetypes=[("Excel files", "*.xlsx")], # Filter to show only Excel files
            title="บันทึกไฟล์ Maximo Data เป็น", # Dialog title
            initialfile=default_file_name # Default file name suggestion now uses the selected sheet name
        )

        if not output_file_path: # If the user clicks Cancel in the save dialog
            messagebox.showinfo("ยกเลิก", "การบันทึกไฟล์ถูกยกเลิก")
            status_label.config(text="การแปลงข้อมูลถูกยกเลิก")
            return

        # Save the new workbook to the chosen path
        new_workbook.save(output_file_path)

        messagebox.showinfo(
            "สำเร็จ",
            f"การแปลงข้อมูลเสร็จสมบูรณ์ และบันทึกเป็นไฟล์ใหม่แล้วที่:\n{output_file_path}"
        )
        status_label.config(text="การแปลงข้อมูลเสร็จสมบูรณ์")

    except FileNotFoundError:
        messagebox.showerror("ข้อผิดพลาด", "ไม่พบไฟล์ Worklist ที่ระบุ กรุณาตรวจสอบเส้นทางไฟล์")
        status_label.config(text="เกิดข้อผิดพลาด: ไม่พบไฟล์")
    except Exception as e:
        messagebox.showerror("เกิดข้อผิดพลาด", f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}\n"
                                               f"โปรดตรวจสอบว่าไฟล์ Excel ไม่ได้ถูกเปิดอยู่")
        status_label.config(text=f"เกิดข้อผิดพลาด: {e}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Excel Worklist Converter")
root.geometry("720x520") # Adjusted height to accommodate the status label
root.resizable(False, False) # Prevent window resizing
root.iconbitmap("./transfer.ico")
root.option_add("*font", "Tahoma 14")

# Variable to store the selected file path (initialized to None)
worklist_file_path = None
# Initialize global variable for lookup data
response_lookup_data = {
    'AB_lookup': {},
    'CD_lookup': {}
}
# Global variable for the selected sheet name
selected_sheet_name = None 
# Global variable for the sheet selection UI elements
sheet_selection_frame = None
sheet_option_menu = None

# Initialize the BooleanVar for highlighting
enable_highlight_var = tk.BooleanVar(value=False) # Default to false (not highlighted)
# Initialize the BooleanVar for including strikethrough rows (default to false, meaning skip them)
include_strikethrough_rows_var = tk.BooleanVar(value=False) 

# Frame for Worklist File Selection
file_frame = tk.LabelFrame(root, text="1. Worklist File", padx=10, pady=10)
file_frame.pack(pady=10, padx=20, fill="x")

# Button to select file
select_file_button = tk.Button(file_frame, text="เลือกไฟล์ Excel", command=select_excel_file)
select_file_button.pack(side=tk.LEFT, padx=(0, 10))

# Label to display the selected file path
file_path_label = tk.Label(file_frame, text="ยังไม่ได้เลือกไฟล์ Worklist", wraplength=350, justify="left")
file_path_label.pack(side=tk.LEFT, fill="x", expand=True)

# Highlighting and Exclusion options frame
options_frame = tk.LabelFrame(root, text="ตัวเลือกผลลัพธ์", padx=10, pady=5)
options_frame.pack(pady=5, padx=20, fill="x")

# Checkbutton for enabling highlighting
highlight_check = tk.Checkbutton(
    options_frame, 
    text="1. เปิดใช้งานการเน้นสี (Column Activity หากตัวอักษรเกิน 85 ตัว)",
    variable=enable_highlight_var,
    bg="orange"
)
highlight_check.pack(side=tk.TOP, anchor="w", pady=2) 

# Checkbutton for enabling including strikethrough rows
strikethrough_check = tk.Checkbutton(
    options_frame,
    text="2. ดึง Row ที่มี strikethrough (ไม่ดึงเป็นค่าเริ่มต้น)",
    variable=include_strikethrough_rows_var
)
strikethrough_check.pack(side=tk.TOP, anchor="w", pady=2)

# Convert to Maximo button (initially disabled until a file is selected)
convert_button = tk.Button(root, text="Convert to Maximo", command=run_conversion_process, state=tk.DISABLED)
convert_button.pack(pady=20)

# Status label to show processing messages
status_label = tk.Label(root, text="", fg="blue", font=("Tahoma", 12))
status_label.pack(pady=5)

# Label for the update information at the bottom right
update_info_label = tk.Label(
    root,
    text="อัปเดตล่าสุด 20/6/2568 โดย นศ.ฝึกงาน ปิยะ",
    font=("Tahoma", 10), # Smaller font size for this info
    fg="gray" # Gray color for less prominence
)
update_info_label.pack(side=tk.BOTTOM, anchor="se", padx=10, pady=5) # Position at bottom-right

# Start the Tkinter event loop
root.mainloop()