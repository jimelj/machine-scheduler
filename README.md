# Machine Scheduler

A web-based application for optimizing insert scheduling across multiple mail insertion machines, designed to maximize efficiency and minimize changeovers for mail processing operations.

## Overview

Machine Scheduler analyzes pick lists from PDF files and creates an optimized schedule for distributing mail inserts across multiple machines. It provides two scheduling methods:

1. **By Store** - Prioritizes keeping all inserts for the same store on the same machine
2. **By Zipcode** - Prioritizes keeping all inserts for the same zipcode on a single machine

The application organizes workloads by mail date (MON, TUES, etc.) and creates an optimal run sequence to minimize machine changeovers.

## Features

- **PDF Processing**: Automatically extracts and parses pick list data from PDF files
- **Mail Date Prioritization**: Processes mail in the correct chronological order (MON → TUES → etc.)
- **Intelligent Scheduling Algorithms**:
  - Store-based scheduling to minimize machine changes for each store
  - Zipcode-based scheduling to keep zip codes on a single machine
  - Workload balancing across machines
  - Optimized run sequences to minimize changeovers
- **Comprehensive Reports**:
  - Web-based visual reports with interactive elements
  - Excel reports with multiple sheets for detailed analysis
  - Daily machine workload visualization
  - Machine load distribution with progress bars
- **Flexible Configuration**: Supports different numbers of machines and scheduling methods

## Screenshots

(Add screenshots here)

## Installation

### Prerequisites

- Python 3.8+
- pip (Python package installer)

### Setup

1. Clone the repository:
   ```
   git clone https://github.com/your-username/machine-scheduler.git
   cd machine-scheduler
   ```

2. Create and activate a virtual environment:
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install required packages:
   ```
   pip install -r requirements.txt
   ```

4. Prepare the Zips by Address File:
   - Place your "Zips by Address File Group.xlsx" in the project root directory
   - This file should contain columns for zipcode and mail date (MON, TUES, etc.)

## Usage

1. Start the application:
   ```
   python app.py
   ```

2. Open your web browser and navigate to:
   ```
   http://127.0.0.1:5000
   ```

3. Upload a PDF pick list file:
   - Select the number of machines (default is 3)
   - Choose a scheduling method (By Store or By Zipcode)
   - Click "Process PDF"

4. View and analyze the results:
   - Machine load distribution
   - Daily machine workload breakdown
   - Daily machine schedule organized by mail date
   - Suggested run sequence showing optimal order of operations
   - Zipcode machine assignments

5. Download the Excel report for offline analysis

## How It Works

### Scheduling Algorithms

#### Store-Based Scheduling

This algorithm focuses on keeping all inserts for the same store on the same machine:

1. Analyzes which stores appear together most frequently across zip codes
2. Creates an initial machine schedule based on store commonality
3. Processes in mail date order, maintaining the day sequence
4. Balances workloads across machines by moving stores as needed

#### Zipcode-Based Scheduling

This algorithm prioritizes keeping zip codes on a single machine:

1. Processes mail dates in sequence (MON, TUES, etc.)
2. For each day, assigns zip codes to machines based on:
   - Insert continuity (maximizing overlap between inserts)
   - Workload balance across machines
3. Ensures each zipcode only appears on exactly one machine

### Run Sequence Optimization

The application creates an optimal run sequence for each machine by:

1. Processing mail dates in order (MON first, TUES next, etc.)
2. Starting with the largest quantity store for each day
3. Finding the next store with maximum insert continuity
4. Continuing this process to minimize changeovers

## Required Files

- **Pick List PDF**: Contains the material pick lists with zipcode, store, and quantity information
- **Zips by Address File**: An Excel/CSV file that maps zip codes to mail dates (MON, TUES, etc.)

## Troubleshooting

- If you encounter indentation errors, ensure your Python environment is properly configured
- If the PDF parsing fails, check that your PDF format matches the expected structure
- For mail date issues, verify your "Zips by Address File Group.xlsx" file contains the correct format

## License

[Add your license information here]

## Contributing

[Add contribution guidelines here]

## Contact

[Add your contact information here] 