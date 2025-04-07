import pandas as pd
import numpy as np
from collections import defaultdict
import PyPDF2
import io
import re
import os
from flask import Flask, render_template, request, send_file, redirect, url_for, flash, session
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "machine_scheduling_app_secret_key"
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size

# Add enumerate and sum functions to Jinja2 environment
app.jinja_env.globals.update(enumerate=enumerate, sum=sum)

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def extract_text_from_pdf(pdf_file):
    """
    Extract text from a PDF file.
    
    Args:
        pdf_file: File object or path to PDF file
    
    Returns:
        str: Extracted text from PDF
    """
    if isinstance(pdf_file, str):
        # If a path is provided
        with open(pdf_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page_num in range(len(pdf_reader.pages)):
                text += pdf_reader.pages[page_num].extract_text()
    else:
        # If a file object is provided
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            text += pdf_reader.pages[page_num].extract_text()
    
    return text

def parse_pick_lists(data):
    """
    Parse the pick list data into a structured format.
    
    Args:
        data (str): Raw pick list data
    
    Returns:
        dict: Structured data with zip codes, inserts, and stores
    """
    # Split the data by "Material Pick List" to get individual zip code sections
    sections = data.split("Material Pick List")
    
    structured_data = {}
    
    for section in sections:
        if "Zipcode - " not in section:
            continue
            
        # Extract zip code
        zipcode_match = re.search(r"Zipcode - (\d+)", section)
        if not zipcode_match:
            continue
            
        zipcode = zipcode_match.group(1)
        
        # Extract inserts count
        inserts_match = re.search(r"Inserts - (\d+)", section)
        num_inserts = int(inserts_match.group(1)) if inserts_match else 0
        
        # Extract stores and quantities
        stores = []
        lines = section.split('\n')
        store_data_section = False
        
        for i, line in enumerate(lines):
            # Check for table header line
            if (("Store" in line and "Qty" in line) or 
                ("Store" in line and "Wght" in line) or 
                ("Store" in line and "Quantity" in line)):
                store_data_section = True
                continue
                
            if store_data_section and ("Page:" in line or line.strip() == "0"):
                break
                
            # Process lines in the store data section
            if store_data_section and line.strip() and "Total -" not in line and "Machine#" not in line and "Day#" not in line:
                # Clean the line - multiple spaces to single space
                clean_line = ' '.join(line.split())
                
                try:
                    # First, handle the format in the screenshot with multiple columns
                    # Look for a pattern of text, followed by a number (quantity), followed by another number (weight)
                    match = re.search(r'^(.*?)\s+(\d{1,3}(?:,\d{3})?)\s+\d+$', clean_line)
                    if match:
                        store_name = match.group(1).strip()
                        quantity = int(match.group(2).replace(',', ''))
                        stores.append({
                            'store_name': store_name,
                            'quantity': quantity
                        })
                        continue
                        
                    # Try alternative pattern: store name followed by quantity at the end
                    match = re.search(r'^(.*?)\s+(\d{1,3}(?:,\d{3})?)$', clean_line)
                    if match:
                        # For patterns like "SHOPRITE BETHPAGE LI Z5 13,550  1186"
                        # Extract store name and quantity
                        raw_store_name = match.group(1).strip()
                        # Remove trailing patterns like "Z5 13,550"
                        store_name = re.sub(r'\s+(?:Z\d+|ABR|[A-Z]-Z\d+|1-Z\d+)\s+\d{1,3}(?:,\d{3})?$', '', raw_store_name).strip()
                        quantity = int(match.group(2).replace(',', ''))
                        stores.append({
                            'store_name': store_name,
                            'quantity': quantity
                        })
                except Exception:
                    # If we couldn't parse the line, just skip it
                    pass
        
        if zipcode and stores:
            structured_data[zipcode] = {
                'num_inserts': num_inserts,
                'stores': stores
            }
    
    return structured_data

def analyze_common_copies(data):
    """
    Analyze which stores/copies appear in multiple zip codes.
    
    Args:
        data (dict): Structured pick list data
    
    Returns:
        dict: Store frequencies across zip codes
    """
    store_frequencies = defaultdict(list)
    
    for zipcode, info in data.items():
        for store in info['stores']:
            store_name = store['store_name']
            store_frequencies[store_name].append({
                'zipcode': zipcode,
                'quantity': store['quantity']
            })
    
    # Sort by frequency (number of zip codes a store appears in)
    sorted_stores = {k: v for k, v in sorted(
        store_frequencies.items(), 
        key=lambda item: len(item[1]), 
        reverse=True
    )}
    
    return sorted_stores

def create_machine_schedule(common_copies, num_machines=3):
    """
    Create a machine schedule based on common copies.
    
    Args:
        common_copies (dict): Store frequencies across zip codes
        num_machines (int): Number of machines available
    
    Returns:
        dict: Machine assignments for each store
    """
    machine_loads = [0] * num_machines
    machine_assignments = {}
    
    # First, assign the most common copies (appear in most zip codes)
    for store, appearances in common_copies.items():
        # Find the machine with the lowest current load
        machine_idx = machine_loads.index(min(machine_loads))
        
        # Assign this store to the machine
        machine_assignments[store] = machine_idx + 1  # 1-indexed machines
        
        # Update the machine load (number of zip codes this copy appears in)
        machine_loads[machine_idx] += len(appearances)
    
    return machine_assignments, machine_loads

def generate_detailed_schedule(data, machine_assignments):
    """
    Generate a detailed schedule for each machine.
    
    Args:
        data (dict): Structured pick list data
        machine_assignments (dict): Machine assignments for each store
    
    Returns:
        dict: Detailed schedule by machine
    """
    machine_schedule = defaultdict(list)
    
    # Create zipcode to machine mapping for easier processing
    zipcode_machine_mapping = defaultdict(set)
    
    for store, machine in machine_assignments.items():
        for zipcode, info in data.items():
            for store_info in info['stores']:
                if store_info['store_name'] == store:
                    zipcode_machine_mapping[zipcode].add(machine)
    
    # For each zip code, list the machine assignments
    zipcode_schedule = {}
    for zipcode, machines in zipcode_machine_mapping.items():
        zipcode_schedule[zipcode] = list(machines)
    
    # For each machine, list the assigned stores and zip codes
    for store, machine in machine_assignments.items():
        zip_appearances = []
        total_quantity = 0
        
        for zipcode, info in data.items():
            for store_info in info['stores']:
                if store_info['store_name'] == store:
                    zip_appearances.append({
                        'zipcode': zipcode,
                        'quantity': store_info['quantity']
                    })
                    total_quantity += store_info['quantity']
        
        machine_schedule[machine].append({
            'store': store,
            'zip_codes': [app['zipcode'] for app in zip_appearances],
            'zip_code_count': len(zip_appearances),
            'total_quantity': total_quantity
        })
    
    # Sort each machine's assignments by total quantity
    for machine in machine_schedule:
        machine_schedule[machine] = sorted(
            machine_schedule[machine],
            key=lambda x: x['total_quantity'],
            reverse=True
        )
    
    return machine_schedule, zipcode_schedule

def create_machine_schedule_by_zipcode(data, num_machines=3):
    """
    Create a machine schedule that prioritizes keeping zip codes on single machines.
    
    Args:
        data (dict): Structured pick list data
        num_machines (int): Number of machines available
    
    Returns:
        dict: Machine assignments for each store
    """
    machine_loads = [0] * num_machines
    machine_assignments = {}
    zipcode_machine = {}
    
    # Sort zip codes by number of stores/quantity
    zipcode_weights = {}
    for zipcode, info in data.items():
        zipcode_weights[zipcode] = sum(store['quantity'] for store in info['stores'])
    
    # Sort zip codes by weight (descending)
    sorted_zipcodes = sorted(zipcode_weights.items(), key=lambda x: x[1], reverse=True)
    
    # First, assign each zip code to a machine
    for zipcode, weight in sorted_zipcodes:
        # Find the machine with the lowest current load
        machine_idx = machine_loads.index(min(machine_loads))
        
        # Assign this zip code to the machine
        zipcode_machine[zipcode] = machine_idx + 1  # 1-indexed machines
        
        # Update the machine load
        machine_loads[machine_idx] += weight
    
    # Now, assign each store to the machine of its zip code
    # If a store appears in multiple zip codes, use the first one's machine
    for zipcode, info in data.items():
        machine = zipcode_machine[zipcode]
        for store in info['stores']:
            store_name = store['store_name']
            if store_name not in machine_assignments:
                machine_assignments[store_name] = machine
    
    return machine_assignments, machine_loads, zipcode_machine

def allowed_file(filename):
    """Check if the file has an allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'

def get_zip_mail_dates():
    """
    Read the mail dates from the Zips by address Excel file.
    
    Returns:
        dict: Mapping of zip codes to mail dates
    """
    try:
        # Try to find the Excel file
        base_dir = os.path.dirname(app.config['UPLOAD_FOLDER'])
        possible_names = [
            "Zips by Address File Group.xlsx",  # Exact file name provided by user
            "Zips by address.xlsx", 
            "Zips by address.xls", 
            "zips by address.xlsx", 
            "zips_by_address.xlsx",
            "Zips_by_address.xlsx",
            "Zips by address.csv",
            "zips_by_address.csv"
        ]
        
        # Log the base directory we're searching in
        print(f"Searching for zip address file in: {base_dir}")
        
        # Look for all Excel/CSV files in the directory
        excel_files = []
        for file in os.listdir(base_dir):
            if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.csv'):
                excel_files.append(file)
        
        print(f"Found Excel/CSV files: {excel_files}")
        
        # Check for the specific file we need
        zips_file_path = None
        for name in possible_names:
            path = os.path.join(base_dir, name)
            if os.path.exists(path):
                zips_file_path = path
                print(f"Found zip address file: {path}")
                break
        
        if not zips_file_path:
            print("Could not find Zips by address file. Checking current directory...")
            # Try looking in the current working directory as fallback
            current_dir = os.getcwd()
            for name in possible_names:
                path = os.path.join(current_dir, name)
                if os.path.exists(path):
                    zips_file_path = path
                    print(f"Found zip address file in current directory: {path}")
                    break
        
        if not zips_file_path:
            print("Could not find Zips by address file in any location")
            return {}
        
        # Read the Excel or CSV file
        if zips_file_path.endswith('.csv'):
            zips_df = pd.read_csv(zips_file_path)
        else:
            zips_df = pd.read_excel(zips_file_path)
            
        print(f"File successfully read with {len(zips_df)} rows")
        print(f"Excel file columns: {list(zips_df.columns)}")
        print(f"First 5 rows of data:\n{zips_df.head().to_string()}")
        
        # Check for column names and normalize them
        columns = zips_df.columns
        zip_col = None
        mail_date_col = None
        
        # Find the zip code and mail date columns
        for col in columns:
            col_lower = str(col).lower()
            if col_lower in ['zip', 'zipcode', 'zip code']:
                zip_col = col
                print(f"Found zip column: {col}")
            elif col_lower in ['mailday', 'mail date', 'mail day', 'maildate']:
                mail_date_col = col
                print(f"Found mail date column: {col}")
        
        # If column names aren't matching, try a more flexible approach
        if not zip_col:
            for col in columns:
                if 'zip' in str(col).lower():
                    zip_col = col
                    print(f"Using zip-like column: {col}")
                    break
        
        if not mail_date_col:
            for col in columns:
                if 'mail' in str(col).lower() or 'day' in str(col).lower() or 'date' in str(col).lower():
                    mail_date_col = col
                    print(f"Using mail-date-like column: {col}")
                    break
        
        # If column detection failed, hard code based on known excel format from screenshot
        if not zip_col and len(columns) >= 1:
            zip_col = columns[0]  # First column is zip
            print(f"Using first column as zip: {zip_col}")
        
        if not mail_date_col and len(columns) >= 3:
            mail_date_col = columns[2]  # Third column is mail date
            print(f"Using third column as mail date: {mail_date_col}")
        
        # If we still couldn't find the needed columns, return empty dict
        if not zip_col or not mail_date_col:
            print("Could not identify required columns in Excel file")
            return {}
        
        # Create a mapping of zip codes to mail dates
        zip_mail_dates = {}
        for idx, row in zips_df.iterrows():
            try:
                zip_code = str(row[zip_col]).strip()
                mail_date = str(row[mail_date_col]).strip()
                
                # Handle NaN values
                if zip_code == 'nan' or mail_date == 'nan':
                    continue
                    
                # Remove any decimal part if the zip code is read as a float
                if '.' in zip_code:
                    zip_code = zip_code.split('.')[0]
                
                # Ensure 5-digit formatting
                zip_code = zip_code.zfill(5)
                zip_mail_dates[zip_code] = mail_date
                
                # Debug output for the first few rows
                if idx < 5:
                    print(f"Row {idx}: Zip={zip_code}, Mail Date={mail_date}")
            except Exception as row_error:
                print(f"Error processing row {idx}: {row_error}")
        
        print(f"Loaded {len(zip_mail_dates)} zip codes with mail dates")
        print(f"Sample zip mail dates: {list(zip_mail_dates.items())[:5]}")
        return zip_mail_dates
    except Exception as e:
        print(f"Error reading Zips by address file: {e}")
        import traceback
        traceback.print_exc()
        return {}

def process_pdf_file(pdf_path, num_machines, scheduling_method='by_store'):
    """Process PDF file and return schedule data"""
    # Extract text from PDF
    raw_data = extract_text_from_pdf(pdf_path)
    
    # Parse the data
    parsed_data = parse_pick_lists(raw_data)
    
    # Get mail dates from the Zips by address file
    mail_dates = get_zip_mail_dates()
    
    # Group zipcodes by mail date
    zipcode_by_date = defaultdict(list)
    for zipcode in parsed_data.keys():
        mail_date = mail_dates.get(zipcode, "")
        zipcode_by_date[mail_date].append(zipcode)
    
    # Print mail date groupings for debugging
    print("Zipcodes by mail date:")
    for date, zipcodes in zipcode_by_date.items():
        print(f"{date}: {zipcodes}")
    
    # Analyze common copies
    common_copies = analyze_common_copies(parsed_data)
    
    # Create machine schedule based on selected method
    if scheduling_method == 'by_zipcode':
        # Use zip code prioritized scheduling for each day separately
        machine_schedule = {}
        for machine_num in range(1, num_machines + 1):
            machine_schedule[machine_num] = []
        
        # Create zipcode schedule
        zipcode_schedule = {}
        
        # Initialize machine_loads for by_zipcode method (will track total quantities)
        machine_loads = [0] * num_machines
        
        # Group data by mail date
        data_by_date = defaultdict(dict)
        for zipcode, info in parsed_data.items():
            mail_date = mail_dates.get(zipcode, "")
            data_by_date[mail_date][zipcode] = info
        
        # Get ordered mail dates (MON first, then TUES, etc.)
        ordered_days = ["MON", "TUES", "WED", "THURS", "FRI", "SAT", "SUN", ""]
        sorted_mail_dates = sorted(data_by_date.keys(), key=lambda x: ordered_days.index(x) if x in ordered_days else len(ordered_days))
        
        # Process each mail date separately
        for mail_date in sorted_mail_dates:
            date_data = data_by_date[mail_date]
            
            if not date_data:
                continue
                
            # Schedule for this day only
            day_assignments, day_loads, day_zipcode_machine = create_machine_schedule_by_zipcode(date_data, num_machines)
            
            # Update zipcode schedule
            for zipcode, machine in day_zipcode_machine.items():
                zipcode_schedule[zipcode] = {
                    'machines': [machine],
                    'mail_date': mail_date
                }
            
            # Add stores to machine assignments for this day
            for store, machine in day_assignments.items():
                # Find zip codes and quantities for this store
                zip_appearances = []
                total_quantity = 0
                
                for zipcode, info in date_data.items():
                    for store_info in info['stores']:
                        if store_info['store_name'] == store:
                            zip_appearances.append({
                                'zipcode': zipcode,
                                'quantity': store_info['quantity'],
                                'mail_date': mail_date
                            })
                            total_quantity += store_info['quantity']
                
                # Add to machine schedule
                machine_schedule[machine].append({
                    'store': store,
                    'zip_codes': [app['zipcode'] for app in zip_appearances],
                    'mail_date': mail_date,
                    'zip_code_count': len(zip_appearances),
                    'total_quantity': total_quantity
                })
                
                # Update total quantity for this machine
                machine_loads[machine-1] += total_quantity
    else:
        # Use store-based scheduling but process days in order
        # First, create separate store appearances by mail date
        store_appearances_by_date = defaultdict(lambda: defaultdict(list))
        
        for zipcode, info in parsed_data.items():
            mail_date = mail_dates.get(zipcode, "")
            
            for store in info['stores']:
                store_name = store['store_name']
                store_appearances_by_date[mail_date][store_name].append({
                    'zipcode': zipcode,
                    'quantity': store['quantity']
                })
        
        # Get ordered mail dates (MON first, then TUES, etc.)
        ordered_days = ["MON", "TUES", "WED", "THURS", "FRI", "SAT", "SUN", ""]
        sorted_mail_dates = sorted(store_appearances_by_date.keys(), key=lambda x: ordered_days.index(x) if x in ordered_days else len(ordered_days))
        
        # Create machine assignments
        machine_assignments, machine_loads = create_machine_schedule(common_copies, num_machines)
        
        # Organize assignments by day first, then by machine
        machine_schedule = {}
        zipcode_machine_mapping = defaultdict(set)
        
        # Initialize machine loads with total quantities
        machine_total_quantities = [0] * num_machines
        
        for machine_num in range(1, num_machines + 1):
            machine_schedule[machine_num] = []
        
        # First pass: assign stores to machines based on common copies
        for mail_date in sorted_mail_dates:
            for store, machine in machine_assignments.items():
                # Skip if this store doesn't have appearances for this day
                if store not in store_appearances_by_date[mail_date]:
                    continue
                    
                store_appearances = store_appearances_by_date[mail_date][store]
                zip_codes = [app['zipcode'] for app in store_appearances]
                total_quantity = sum(app['quantity'] for app in store_appearances)
                
                if zip_codes:
                    # Update zipcode machine mapping
                    for zipcode in zip_codes:
                        zipcode_machine_mapping[zipcode].add(machine)
                    
                    # Add to machine schedule
                    machine_schedule[machine].append({
                        'store': store,
                        'zip_codes': zip_codes,
                        'mail_date': mail_date,
                        'zip_code_count': len(zip_codes),
                        'total_quantity': total_quantity
                    })
                    
                    # Update total quantity for this machine
                    machine_total_quantities[machine-1] += total_quantity
        
        # Create zipcode schedule with mail dates
        zipcode_schedule = {}
        for zipcode, machines in zipcode_machine_mapping.items():
            mail_date = mail_dates.get(zipcode, "")
            zipcode_schedule[zipcode] = {
                'machines': list(machines),
                'mail_date': mail_date
            }
        
        # Use the total quantities instead of zip counts for machine_loads
        machine_loads = machine_total_quantities
    
    # Calculate machine loads per day (zip code counts)
    machine_loads_by_date = defaultdict(lambda: [0] * num_machines)
    for machine, assignments in machine_schedule.items():
        for assignment in assignments:
            mail_date = assignment.get('mail_date', '')
            machine_loads_by_date[mail_date][machine-1] += assignment['zip_code_count']
    
    # Convert to regular dict for JSON serialization
    machine_loads_by_date = {date: loads for date, loads in machine_loads_by_date.items()}
    
    # Calculate total load (total quantities)
    total_load = sum(machine_loads)
    
    return {
        'machine_schedule': machine_schedule,
        'zipcode_schedule': zipcode_schedule,
        'machine_loads': machine_loads,
        'machine_loads_by_date': machine_loads_by_date,
        'total_load': total_load,
        'zip_code_count': len(parsed_data),
        'mail_dates': sorted_mail_dates
    }

def create_excel_report(machine_schedule, pdf_path, zipcode_schedule=None, machine_loads_by_date=None, mail_dates=None):
    """Create Excel report from schedule data"""
    # Create Excel writer with multiple sheets
    output_dir = os.path.dirname(os.path.abspath(__file__))  # Save to the app root directory
    excel_path = os.path.join(output_dir, "machine_schedule.xlsx")
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    
    # Get all mail dates
    if not mail_dates:
        # Extract mail dates from assignments if not provided
        all_mail_dates = set()
        for machine, assignments in machine_schedule.items():
            for assignment in assignments:
                mail_date = assignment.get('mail_date', '')
                if mail_date:
                    all_mail_dates.add(mail_date)
        mail_dates = sorted(list(all_mail_dates))
    
    # First sheet - machine schedules organized by mail date
    schedule_df = pd.DataFrame(columns=['Mail Date', 'Machine', 'Store', 'Zip Codes', 'Quantity'])
    
    row_idx = 0
    for mail_date in mail_dates:
        if not mail_date:  # Skip empty mail dates
            continue
            
        for machine, assignments in machine_schedule.items():
            # Filter assignments for this mail date
            date_assignments = [a for a in assignments if a.get('mail_date', '') == mail_date]
            
            for assignment in date_assignments:
                zip_count = assignment['zip_code_count']
                schedule_df.loc[row_idx] = [
                    mail_date,
                    f"Machine {machine}",
                    assignment['store'],
                    f"{', '.join(assignment['zip_codes'])} {zip_count}",
                    assignment['total_quantity']
                ]
                row_idx += 1
    
    # Add any assignments without mail dates at the end
    for machine, assignments in machine_schedule.items():
        # Filter assignments with no mail date
        no_date_assignments = [a for a in assignments if not a.get('mail_date', '')]
        
        for assignment in no_date_assignments:
            zip_count = assignment['zip_code_count']
            schedule_df.loc[row_idx] = [
                "UNASSIGNED",
                f"Machine {machine}",
                assignment['store'],
                f"{', '.join(assignment['zip_codes'])} {zip_count}",
                assignment['total_quantity']
            ]
            row_idx += 1
    
    # Write the machine schedule sheet
    schedule_df.to_excel(writer, sheet_name='Machine Schedule', index=False)
    
    # Second sheet - separate sheets for each mail date
    for mail_date in mail_dates:
        if not mail_date:  # Skip empty mail dates
            continue
            
        date_df = pd.DataFrame(columns=['Machine', 'Store', 'Zip Codes', 'Quantity'])
        row_idx = 0
        
        for machine, assignments in machine_schedule.items():
            # Filter assignments for this mail date
            date_assignments = [a for a in assignments if a.get('mail_date', '') == mail_date]
            
            for assignment in date_assignments:
                zip_count = assignment['zip_code_count']
                date_df.loc[row_idx] = [
                    f"Machine {machine}",
                    assignment['store'],
                    f"{', '.join(assignment['zip_codes'])} {zip_count}",
                    assignment['total_quantity']
                ]
                row_idx += 1
        
        if not date_df.empty:
            date_df.to_excel(writer, sheet_name=f'{mail_date} Schedule', index=False)
    
    # Third sheet - zipcode to machine mapping with mail dates
    if zipcode_schedule:
        print("Creating zipcode schedule sheet with mail dates")
        zipcode_df = pd.DataFrame(columns=['Mail Date', 'Zipcode', 'Machines'])
        
        # Group zipcodes by mail date
        by_date = defaultdict(list)
        for zipcode, data in zipcode_schedule.items():
            mail_date = data.get('mail_date', 'UNASSIGNED')
            if not mail_date:
                mail_date = 'UNASSIGNED'
            by_date[mail_date].append((zipcode, data))
        
        # Add to dataframe in order of mail date
        row_idx = 0
        for mail_date in mail_dates + ['UNASSIGNED']:
            for zipcode, data in sorted(by_date.get(mail_date, [])):
                machines = ", ".join(map(str, data.get('machines', [])))
                zipcode_df.loc[row_idx] = [
                    mail_date,
                    zipcode,
                    machines
                ]
                row_idx += 1
        
        # Debug: Print the entire dataframe before writing to Excel
        print(f"Zipcode DataFrame:\n{zipcode_df.head(10).to_string()}")
        
        # Write the zipcode sheet
        zipcode_df.to_excel(writer, sheet_name='Zipcode Schedule', index=False)
    
    # Fourth sheet - Machine loads by day
    if machine_loads_by_date:
        loads_df = pd.DataFrame(columns=['Mail Date'] + [f'Machine {i+1}' for i in range(len(next(iter(machine_loads_by_date.values()))))])
        
        row_idx = 0
        for mail_date in mail_dates:
            if mail_date in machine_loads_by_date:
                loads = machine_loads_by_date[mail_date]
                loads_df.loc[row_idx] = [mail_date] + loads
                row_idx += 1
        
        if not loads_df.empty:
            loads_df.to_excel(writer, sheet_name='Daily Machine Loads', index=False)
    
    # Save the Excel file
    writer.close()
    print(f"Excel report saved to: {excel_path}")
    
    return excel_path

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'pdf_file' not in request.files:
            flash('No file part', 'danger')
            return redirect(request.url)
            
        pdf_file = request.files['pdf_file']
        
        if pdf_file.filename == '':
            flash('No selected file', 'danger')
            return redirect(request.url)
            
        if pdf_file and allowed_file(pdf_file.filename):
            try:
                # Get the number of machines
                num_machines = int(request.form.get('machines', 3))
                
                # Get scheduling method
                scheduling_method = request.form.get('scheduling_method', 'by_store')
                
                # Save the uploaded PDF file
                pdf_filename = secure_filename(pdf_file.filename)
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
                pdf_file.save(pdf_path)
                
                # Process the PDF file
                result = process_pdf_file(pdf_path, num_machines, scheduling_method)
                
                # Generate Excel report
                excel_path = create_excel_report(
                    machine_schedule=result['machine_schedule'], 
                    pdf_path=pdf_path, 
                    zipcode_schedule=result.get('zipcode_schedule'),
                    machine_loads_by_date=result.get('machine_loads_by_date', {}),
                    mail_dates=result.get('mail_dates', [])
                )
                
                # Verify that the Excel file was created successfully
                if not os.path.exists(excel_path):
                    raise FileNotFoundError(f"Failed to create Excel file at {excel_path}")
                    
                excel_filename = os.path.basename(excel_path)
                
                # Add default values for variables that might be missing
                machine_loads = result.get('machine_loads', [0] * num_machines)
                total_load = result.get('total_load', 0)
                zip_code_count = result.get('zip_code_count', 0)
                
                # Define all required variables for the template
                template_vars = {
                    'machine_schedule': result.get('machine_schedule', {}),
                    'zipcode_schedule': result.get('zipcode_schedule', {}),
                    'machine_loads': machine_loads,
                    'total_load': total_load,
                    'zip_code_count': zip_code_count,
                    'excel_filename': excel_filename,
                    'machine_loads_by_date': result.get('machine_loads_by_date', {}),
                    'mail_dates': result.get('mail_dates', [])
                }
                
                return render_template('results.html', **template_vars)
                
            except Exception as e:
                flash(f'Error processing file: {str(e)}', 'danger')
                import traceback
                traceback.print_exc()
                return redirect(request.url)
        else:
            flash('Allowed file types are PDF only', 'danger')
            return redirect(request.url)
            
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    """Download the Excel report"""
    # Use the application's root directory to find the file
    app_root = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(app_root, filename)
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash(f"File not found: {filename}", "danger")
        return redirect(url_for('index'))

# Create a simple error handler
@app.errorhandler(Exception)
def handle_error(e):
    return render_template('error.html', error=str(e)), 500

if __name__ == '__main__':
    app.run(debug=True)