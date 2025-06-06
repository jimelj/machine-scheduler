import pandas as pd
import numpy as np
from collections import defaultdict, Counter
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
    Analyze which stores appear together most frequently across zip codes.
    Returns a dictionary of store pairs to their frequency.
    """
    common_copies = defaultdict(int)
    
    # First, get all unique stores across all zip codes
    all_stores = set()
    for zipcode, info in data.items():
        for store in info['stores']:
            all_stores.add(store['store_name'])
    
    # Create a mapping of zip codes to store names for easier lookup
    zipcode_stores = {}
    for zipcode, info in data.items():
        zipcode_stores[zipcode] = [store['store_name'] for store in info['stores']]
    
    # Now analyze which stores appear together
    for zipcode, store_names in zipcode_stores.items():
        # For each pair of stores in this zip code
        for i, store1 in enumerate(store_names):
            for store2 in store_names[i+1:]:
                # Order the store names to ensure consistent pairing
                pair = tuple(sorted([store1, store2]))
                common_copies[pair] += 1
    
    # Convert to regular dict for return
    return dict(common_copies)

def create_machine_schedule(common_copies, num_machines):
    """
    Create a machine schedule optimized for minimizing changeovers and balancing workloads.
    
    Args:
        common_copies: Dictionary of store pairs to their frequency
        num_machines: Number of machines available
    
    Returns:
        machine_assignments: Dictionary mapping stores to machines
        machine_loads: List of machine loads
    """
    # Sort store pairs by frequency in descending order
    sorted_pairs = sorted(common_copies.items(), key=lambda x: x[1], reverse=True)
    
    # Create machine assignments and load tracking
    machine_assignments = {}
    machine_stores = defaultdict(set)  # Stores assigned to each machine
    machine_loads = [0] * num_machines
    
    # Initialize a graph of store relationships
    store_graph = defaultdict(set)
    for (store1, store2), _ in sorted_pairs:
        store_graph[store1].add(store2)
        store_graph[store2].add(store1)
    
    # Get unique stores from all pairs
    all_stores = set()
    for store1, store2 in common_copies.keys():
        all_stores.add(store1)
        all_stores.add(store2)
    
    # Calculate a "commonality score" for each store
    # This represents how frequently a store appears with other stores
    store_commonality = {}
    for store in all_stores:
        connected_stores = store_graph[store]
        score = sum(common_copies.get(tuple(sorted([store, other])), 0) for other in connected_stores)
        store_commonality[store] = score
    
    # Sort stores by commonality score (most connected first)
    sorted_stores = sorted(store_commonality.items(), key=lambda x: x[1], reverse=True)
    
    # First, assign highly connected stores to machines evenly
    for store, _ in sorted_stores:
        if store in machine_assignments:
            continue
            
        # Find the machine with the lowest load and assign this store to it
        min_load_machine = machine_loads.index(min(machine_loads))
        machine_assignments[store] = min_load_machine + 1  # 1-indexed for machines
        machine_stores[min_load_machine + 1].add(store)
        machine_loads[min_load_machine] += 1
        
        # Now try to assign related stores to the same machine
        related_stores = list(store_graph[store])
        related_stores.sort(key=lambda s: common_copies.get(tuple(sorted([store, s])), 0), reverse=True)
        
        for related_store in related_stores:
            if related_store in machine_assignments:
                continue
                
            # Check if adding this store would overload the machine
            # More strict load balancing - lower the threshold to 1.2
            if machine_loads[min_load_machine] >= (sum(machine_loads) / num_machines) * 1.2:
                break  # Stop adding related stores if this machine is getting too full
                
            machine_assignments[related_store] = min_load_machine + 1
            machine_stores[min_load_machine + 1].add(related_store)
            machine_loads[min_load_machine] += 1
    
    # Assign any remaining stores
    for store in all_stores:
        if store not in machine_assignments:
            min_load_machine = machine_loads.index(min(machine_loads))
            machine_assignments[store] = min_load_machine + 1
            machine_stores[min_load_machine + 1].add(store)
            machine_loads[min_load_machine] += 1
    
    print(f"Machine assignments: {machine_assignments}")
    print(f"Machine loads: {machine_loads}")
    
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

def create_machine_schedule_by_zipcode(data, num_machines):
    """
    Create a machine schedule prioritizing keeping zip codes on a single machine,
    while also minimizing changeovers by maximizing insert continuity and ensuring
    balanced workloads across machines.
    
    Args:
        data: Dictionary of zip codes to their data
        num_machines: Number of machines available
        
    Returns:
        store_machine: Dictionary mapping stores to machines
        machine_loads: List of machine loads
        zipcode_machine: Dictionary mapping zip codes to machines
    """
    # First, calculate the weight of each zip code
    # Weight is based on total quantity and number of stores
    zipcode_weights = {}
    for zipcode, info in data.items():
        total_quantity = 0
        for store in info['stores']:
            total_quantity += store['quantity']
        zipcode_weights[zipcode] = total_quantity
    
    # Sort zip codes by weight (highest first)
    sorted_zipcodes = sorted(zipcode_weights.items(), key=lambda x: x[1], reverse=True)
    
    # Track which inserts are on each machine
    machine_inserts = defaultdict(set)
    
    # Track the load of each machine
    machine_loads = [0] * num_machines
    
    # Map zipcodes to machines
    zipcode_machine = {}
    
    # Now assign zip codes to machines, prioritizing continuity and load balancing
    for zipcode, weight in sorted_zipcodes:
        # Get the inserts for this zip code
        zipcode_inserts = {store['store_name'] for store in data[zipcode]['stores']}
        
        # Calculate overlap score with each machine's current inserts
        overlap_scores = []
        for machine_num in range(1, num_machines + 1):
            current_inserts = machine_inserts[machine_num]
            
            # Calculate current machine load percentage
            total_load = sum(machine_loads)
            if total_load == 0:
                load_percentage = 0
            else:
                load_percentage = machine_loads[machine_num-1] / total_load
            
            # Load balancing factor (penalize machines that already have high loads)
            load_balance_penalty = load_percentage * 5  # Increased penalty for unbalanced loads
            
            if not current_inserts:  # If machine is empty, score is 0
                overlap_scores.append((machine_num, 0 - load_balance_penalty, machine_loads[machine_num-1]))
                continue
                
            # Calculate overlap (common inserts)
            overlap = len(zipcode_inserts.intersection(current_inserts))
            
            # Calculate how many inserts need to be added (changeover cost)
            new_inserts = len(zipcode_inserts - current_inserts)
            
            # Score prioritizes high overlap, low changeover, and load balancing
            score = overlap - (new_inserts * 0.5) - load_balance_penalty
            overlap_scores.append((machine_num, score, machine_loads[machine_num-1]))
        
        # Sort by score (higher is better) and then by load (lower is better)
        overlap_scores.sort(key=lambda x: (-x[1], x[2]))
        
        # Assign to best machine
        best_machine = overlap_scores[0][0]
        zipcode_machine[zipcode] = best_machine
        
        # Update machine inserts and load
        machine_inserts[best_machine].update(zipcode_inserts)
        machine_loads[best_machine-1] += zipcode_weights[zipcode]
    
    # Create a mapping of stores to the zipcodes they appear in
    store_zipcodes = defaultdict(list)
    for zipcode, info in data.items():
        for store in info['stores']:
            store_name = store['store_name']
            store_zipcodes[store_name].append((zipcode, store['quantity']))
    
    # Now assign stores based on zipcode assignments
    # Ensure that stores are assigned to machines where their zipcodes are assigned
    store_machine = {}
    
    # First, create a mapping of which machine each store appears most on based on zipcode assignments
    store_machine_appearances = defaultdict(lambda: defaultdict(int))
    for store_name, zipcode_appearances in store_zipcodes.items():
        for zipcode, quantity in zipcode_appearances:
            if zipcode in zipcode_machine:
                machine = zipcode_machine[zipcode]
                # Weight by quantity to prioritize higher volume assignments
                store_machine_appearances[store_name][machine] += quantity
    
    # Assign each store to the machine where it has the most zipcode quantity
    for store_name, machine_quantities in store_machine_appearances.items():
        if not machine_quantities:
            # Find the machine with the minimum load if no zipcodes found
            min_load_machine = machine_loads.index(min(machine_loads)) + 1
            store_machine[store_name] = min_load_machine
        else:
            # Assign to the machine with the highest quantity for this store
            best_machine = max(machine_quantities.items(), key=lambda x: x[1])[0]
            store_machine[store_name] = best_machine
    
    print(f"Machine loads: {machine_loads}")
    print(f"Store assignments: {len(store_machine)}")
    print(f"Zipcode assignments: {len(zipcode_machine)}")
    
    return store_machine, machine_loads, zipcode_machine

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
        
        # Create zipcode schedule - this will ensure each zipcode is on exactly one machine
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
            
            # Update zipcode schedule - ensure each zipcode is associated with exactly one machine
            for zipcode, machine in day_zipcode_machine.items():
                zipcode_schedule[zipcode] = {
                    'machines': [machine],  # Single machine array
                    'mail_date': mail_date
                }
            
            # Add stores to machine assignments for this day
            # First, create a mapping of stores to their zip codes for this day
            store_to_zipcodes = defaultdict(list)
            for zipcode, info in date_data.items():
                machine = day_zipcode_machine[zipcode]  # Get assigned machine for this zipcode
                for store_info in info['stores']:
                    store_name = store_info['store_name']
                    if day_assignments.get(store_name) == machine:  # Only include if store is assigned to this zipcode's machine
                        store_to_zipcodes[store_name].append({
                            'zipcode': zipcode,
                            'quantity': store_info['quantity'],
                            'mail_date': mail_date
                        })
            
            # Now create machine assignments only including zipcodes that match the store's machine
            for store, machine in day_assignments.items():
                # Only get zip codes for this store that are assigned to this machine
                zip_appearances = store_to_zipcodes.get(store, [])
                
                if not zip_appearances:
                    continue  # Skip if no matching zipcodes
                
                total_quantity = sum(app['quantity'] for app in zip_appearances)
                
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
        
        # Step 1: Calculate total quantities for each store to get a better estimate of workload
        store_total_quantities = {}
        for mail_date in sorted_mail_dates:
            for store_name, appearances in store_appearances_by_date[mail_date].items():
                if store_name not in store_total_quantities:
                    store_total_quantities[store_name] = 0
                store_total_quantities[store_name] += sum(app['quantity'] for app in appearances)
        
        # Sort stores by quantity for more balanced assignment
        sorted_stores = sorted(store_total_quantities.items(), key=lambda x: x[1], reverse=True)
        
        # Rebalance machine assignments based on actual quantities
        rebalanced_assignments = {}
        current_machine_loads = [0] * num_machines
        
        # Assign highest quantity stores first, alternating between machines
        for store_name, quantity in sorted_stores:
            min_load_machine = current_machine_loads.index(min(current_machine_loads)) + 1  # 1-indexed
            rebalanced_assignments[store_name] = min_load_machine
            current_machine_loads[min_load_machine - 1] += quantity
        
        # Use the rebalanced assignments
        machine_assignments = rebalanced_assignments
        
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
        
        # Check for balance - if too unbalanced, move stores to balance loads
        total_quantity = sum(machine_total_quantities)
        avg_quantity = total_quantity / num_machines
        
        print(f"Initial machine loads: {machine_total_quantities}")
        print(f"Average expected load: {avg_quantity}")
        
        # If any machine is more than 30% above average, try to rebalance
        max_load = max(machine_total_quantities)
        min_load = min(machine_total_quantities)
        
        if max_load > avg_quantity * 1.3:
            print("Loads are unbalanced, attempting to rebalance...")
            
            # Find which machine is overloaded
            overloaded_machines = [i+1 for i, load in enumerate(machine_total_quantities) if load > avg_quantity * 1.3]
            underloaded_machines = [i+1 for i, load in enumerate(machine_total_quantities) if load < avg_quantity * 0.7]
            
            for overloaded_machine in overloaded_machines:
                # Sort stores on this machine by quantity (largest first) to consider moving
                stores_to_consider = []
                for store, machine in machine_assignments.items():
                    if machine == overloaded_machine:
                        total_store_quantity = store_total_quantities.get(store, 0)
                        stores_to_consider.append((store, total_store_quantity))
                
                stores_to_consider.sort(key=lambda x: x[1], reverse=True)
                
                # Try moving some stores to underloaded machines
                for store, quantity in stores_to_consider:
                    # Only move if it helps balance without overloading the destination
                    for target_machine in underloaded_machines:
                        new_src_load = machine_total_quantities[overloaded_machine-1] - quantity
                        new_dest_load = machine_total_quantities[target_machine-1] + quantity
                        
                        # Check if this move improves balance
                        if new_src_load >= new_dest_load * 0.7 and new_src_load <= new_dest_load * 1.3:
                            print(f"Moving store {store} from machine {overloaded_machine} to {target_machine}")
                            # Update machine assignment
                            machine_assignments[store] = target_machine
                            # Update loads
                            machine_total_quantities[overloaded_machine-1] = new_src_load
                            machine_total_quantities[target_machine-1] = new_dest_load
                            break
                    
                    # Check if we've achieved balance
                    if max(machine_total_quantities) <= avg_quantity * 1.3:
                        break
        
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
    
    # STEP: Sort assignments on each machine to optimize run sequence
    # Group assignments by machine and mail date
    assignments_by_machine_date = {}
    for machine_num in range(1, num_machines + 1):
        assignments_by_machine_date[machine_num] = defaultdict(list)
        for assignment in machine_schedule[machine_num]:
            mail_date = assignment.get('mail_date', '')
            assignments_by_machine_date[machine_num][mail_date].append(assignment)
    
    # Function to calculate overlap between two store assignments
    def calculate_store_overlap(store1, store2):
        # Get set of inserts for each store (we use store names as proxies for inserts)
        inserts1 = {store1['store']}
        inserts2 = {store2['store']}
        
        # Calculate overlap
        overlap = len(inserts1.intersection(inserts2))
        
        # Calculate change cost (penalty for changes)
        change_cost = len(inserts1.symmetric_difference(inserts2))
        
        return overlap - change_cost * 0.5
    
    # Sort assignments for each machine and day to optimize continuity
    for machine_num in range(1, num_machines + 1):
        sorted_assignments = []
        
        # First process by mail date (keep mail date ordering)
        for mail_date in sorted_mail_dates:
            if mail_date in assignments_by_machine_date[machine_num]:
                assignments = assignments_by_machine_date[machine_num][mail_date]
                
                # Skip if there's only one assignment
                if len(assignments) <= 1:
                    sorted_assignments.extend(assignments)
                    continue
                
                # Start with the assignment with most inserts
                sorted_date_assignments = []
                remaining = assignments.copy()
                
                # Start with the store that has the highest quantity
                current = max(remaining, key=lambda x: x['total_quantity'])
                sorted_date_assignments.append(current)
                remaining.remove(current)
                
                # Then repeatedly find the next store with highest overlap
                while remaining:
                    current = sorted_date_assignments[-1]
                    next_store = max(remaining, key=lambda x: calculate_store_overlap(current, x))
                    sorted_date_assignments.append(next_store)
                    remaining.remove(next_store)
                
                sorted_assignments.extend(sorted_date_assignments)
        
        # Replace the original assignments with the sorted ones
        machine_schedule[machine_num] = sorted_assignments
    
    # For zipcode method, double-check that each machine's assignments only contain zipcodes assigned to that machine
    if scheduling_method == 'by_zipcode':
        # Create a mapping of which zipcodes belong to which machine
        machine_zipcodes = defaultdict(set)
        for zipcode, data in zipcode_schedule.items():
            if data.get('machines') and len(data['machines']) > 0:
                machine_zipcodes[data['machines'][0]].add(zipcode)
        
        # Now filter each machine's assignments to only include zipcodes that belong to that machine
        for machine_num in range(1, num_machines + 1):
            allowed_zipcodes = machine_zipcodes.get(machine_num, set())
            filtered_assignments = []
            
            for assignment in machine_schedule[machine_num]:
                # Filter the zip codes list to only include those assigned to this machine
                filtered_zip_codes = [z for z in assignment['zip_codes'] if z in allowed_zipcodes]
                
                # Only include this assignment if it has at least one zipcode on this machine
                if filtered_zip_codes:
                    # Create a copy of the assignment with filtered zipcodes
                    filtered_assignment = assignment.copy()
                    filtered_assignment['zip_codes'] = filtered_zip_codes
                    filtered_assignment['zip_code_count'] = len(filtered_zip_codes)
                    filtered_assignments.append(filtered_assignment)
            
            # Replace the machine's assignments with the filtered version
            machine_schedule[machine_num] = filtered_assignments
    
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
    
    # Fifth sheet - Optimized Run Sequence
    # Create a sheet showing the optimal run sequence for each machine
    run_sequence_df = pd.DataFrame(columns=['Run Order', 'Machine', 'Mail Date', 'Store', 'Zip Codes', 'Quantity'])
    
    row_idx = 0
    for machine_num in range(1, len(next(iter(machine_loads_by_date.values()))) + 1):
        run_order = 1
        for assignment in machine_schedule[machine_num]:
            mail_date = assignment.get('mail_date', 'UNASSIGNED')
            zip_count = assignment['zip_code_count']
            
            # Get the zip codes as a readable string
            zip_codes_str = f"{', '.join(assignment['zip_codes'])}"
            if zip_count > 0:
                zip_codes_str += f" ({zip_count})"
            
            run_sequence_df.loc[row_idx] = [
                run_order,
                f"Machine {machine_num}",
                mail_date,
                assignment['store'],
                zip_codes_str,
                assignment['total_quantity']
            ]
            row_idx += 1
            run_order += 1
            
            # Add a blank row between machines
            if run_order > len(machine_schedule[machine_num]):
                run_sequence_df.loc[row_idx] = [
                    "", 
                    "", 
                    "", 
                    "", 
                    "", 
                    ""
                ]
                row_idx += 1
    
    # Write the run sequence sheet
    run_sequence_df.to_excel(writer, sheet_name='Run Sequence', index=False)
    
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
    app.run(host='0.0.0.0', port=5001, debug=True)