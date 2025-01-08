import os
import pdfplumber
import pandas as pd
import csv
import openpyxl
from datetime import datetime, timedelta

def extract_rows_in_range(input_file, start_marker, end_marker):
    extracted_rows = []  # List to store extracted rows
    start_extraction = False  # Flag to indicate when to start extraction

    with open(input_file, 'r') as infile:
        reader = csv.reader(infile)
        
        for row in reader:
            # Check if the current row contains the start marker
            if row and start_marker in row[0]:
                start_extraction = True  # Start extraction from this point
                continue  # Skip the start marker row
            
            # If we have started extraction, append the row to the list
            if start_extraction:
                # Check for the end marker to stop extraction
                if row and end_marker in row[0]:
                    break  # Stop extracting when we hit the end marker
                extracted_rows.append(row)
    return extracted_rows

def extract_rows_in_rangeMed(input_file, start_marker, end_marker):
    extracted_rows = []  # List to store extracted rows
    start_extraction = False  # Flag to indicate when to start extraction

    with open(input_file, 'r') as infile:
        reader = csv.reader(infile)
        
        for row in reader:
            # Check if the current row contains the start marker
            if row and start_marker in ' '.join(row):
                start_extraction = True  # Start extraction from this point
                continue  # Skip the start marker row
            
            # If we have started extraction, append the row to the list
            if start_extraction:
                # Check for the end marker to stop extraction
                if row and end_marker in ' '.join(row):
                    break  # Stop extracting when we hit the end marker
                extracted_rows.append(row)
    
    return extracted_rows

def extract_rows_in_range_xlsx(input_file, start_marker, end_marker):
    # Load the workbook and select the active sheet
    workbook = openpyxl.load_workbook(input_file)
    sheet = workbook.active

    extracted_rows = []
    start_extraction = False

    # Iterate over the rows in the sheet
    for row in sheet.iter_rows(values_only=True):
        row_data = [str(cell) if cell is not None else '' for cell in row]  # Convert cells to strings
        
        # Check for the start marker
        if start_marker in ' '.join(row_data):
            start_extraction = True
            continue  # Skip the start marker row

        # Start extracting rows
        if start_extraction:
            if end_marker in ' '.join(row_data):
                break  # Stop extraction when we hit the end marker
            extracted_rows.append(row_data)

    return extracted_rows

def extract_date_from_filename(filename):
    parts = filename.split()
    day = parts[-3]  # "13"
    month = parts[-2]  # "SEP"
    year = parts[-1].replace(".xlsx", "")  # "24"

    if len(year) == 2:
        year = "20" + year
    
    # Combine and convert to a datetime object
    date_str = f"{day} {month} {year}"
    date_obj = datetime.strptime(date_str, "%d %b %Y")  # Converts to datetime
    return date_obj

def is_valid_date(date_str):
    try:
        # Attempt to parse the string as a date
        datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        return True
    except ValueError:
        # If a ValueError is raised, the string is not a valid date
        return False

def ReadReport(report):
    with pdfplumber.open(report + ".pdf") as pdf:
        all_data = []
        
        # Iterate over each page
        for page in pdf.pages:
            # Extract table-like content from the page
            tables = page.extract_tables()
            
            for table in tables:
                df = pd.DataFrame(table)  # Convert to pandas DataFrame for easier manipulation
                all_data.append(df)

        # Combine data from all pages
        full_data = pd.concat(all_data)

    # Save to a CSV file for easier viewing
    full_data.to_csv(report + ".csv", index=False)
    return report + ".csv"


def ReadReportMED(report):
    with pdfplumber.open(report + ".pdf") as pdf:
        
        with open(report + ".csv", mode='w', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)

            for page in pdf.pages:
            # Extract the text from the page
                text = page.extract_text()

                # Check if text was extracted
                if text:
                    # Split the text into lines
                    lines = text.split('\n')

                    # Write each line to the CSV
                    for line in lines:
                        # Split the line into columns based on spaces or tabs
                        columns = line.split()  # Adjust the delimiter as needed
                        csv_writer.writerow(columns)

    return report + ".csv"


def ParseOutputsTSA(input):
    input_file = input
    output_file = 'filtered_output.csv'

    file_exists = os.path.isfile(output_file)

    direction = ""

    with open(input_file, 'r') as infile, open(output_file, 'a', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        if file_exists:
            writer.writerow([])
            writer.writerow(['Report - ' + input])
            writer.writerow([])
        else:
            writer.writerow(['Report - ' + input])
            writer.writerow([])


        for row in reader:
            # Convert list to string for easier search
            row_str = ''.join(row).strip()

            split_row = row_str.split()

            if "Container" in split_row:
                index = split_row.index("Container")
                GateDir = split_row[index - 1]

                if GateDir == 'Date':
                    writer.writerow(['Gate In Report'])
                    writer.writerow(['Customer', 'Date In', 'Container', 'Job No.'])
                    direction = "in"
                elif GateDir == 'Out':
                    writer.writerow(['Gate Out Report'])
                    writer.writerow(['Customer', 'Date Out', 'Container', 'Release No.'])
                    direction = "out"

            # Check if the row starts with "TNL OTHER" and extract relevant information
            elif row_str.startswith("TNL OTHER"):
                # Split the row into individual pieces
                row_data = row_str.split()

                # Extract customer, date, and container number
                customer = row_data[0] + ' ' + row_data[1]  # Customer
                date = row_data[2]  # Date
                container_number = row_data[4]  # Container number
                job_rel_no = row_data[9]

                # Write the filtered data into the new CSV file
                writer.writerow([customer, date, container_number, job_rel_no])

            elif row_str.startswith("TNL"):
                # Split the row into individual pieces
                row_data = row_str.split()

                # Extract customer, date, and container number
                customer = row_data[0]  # Customer
                date = row_data[1]  # Date
                container_number = row_data[3]  # Container number

                if direction == 'in':
                    job_rel_no = row_data[9]
                    writer.writerow([customer, date, container_number, job_rel_no])

                elif direction == 'out':
                    job_rel_no = row_data[8]
                    writer.writerow([customer, date, container_number, job_rel_no])

def ParseOutputsWTS(input):
    input_file = input
    output_file = 'filtered_output.csv'

    file_exists = os.path.isfile(output_file)

    customer = ""

    extracted_data_title = extract_rows_in_range(input_file, "Container Daily Log", "MANIFEST ADVICES")
    extracted_data_in = extract_rows_in_range(input_file, "GATE IN", "ESTIMATES APPROVED")
    extracted_data_out = extract_rows_in_range(input_file, "GATE OUT", '"GATE OUT REVERSAL')

    with open(input_file, 'r') as infile, open(output_file, 'a', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        if file_exists:
            writer.writerow([])
            writer.writerow(['Report - ' + input])
            writer.writerow([])
        else:
            writer.writerow(['Report - ' + input])
            writer.writerow([])

        writer.writerow(['Gate In'])
        for line in extracted_data_title:
            comps = line[0].split()
            customer = comps[0].replace("CUSTOMER:", "")

        writer.writerow(['Customer', 'Date In', 'Container', 'Job No.'])
        if len(extracted_data_in) > 1:
            for line in extracted_data_in:
                comps = line[0].split()
                if len(comps) > 0 and comps[2] == '20':

                    date_position = next(i for i, v in enumerate(comps) if '/' in v)

                    writer.writerow([customer, comps[date_position], comps[0], comps[date_position + 1]])
        else:
            writer.writerow(['No tanks in.'])

        writer.writerow([])
        writer.writerow(['Gate Out'])
        writer.writerow(['Customer', 'Date Out', 'Container', 'Job No.'])
        if len(extracted_data_out) > 1:
            for line in extracted_data_out:
                comps = line[0].split()
                if len(comps) > 0 and comps[1] == '20':

                    date_position = next(i for i, v in enumerate(comps) if '/' in v)

                    writer.writerow([customer, comps[date_position], comps[0], comps[date_position + 1]])
        else:
            writer.writerow(['No tanks out.'])

def ParseOutputsMED(input):
    input_file = input
    output_file = 'filtered_output.csv'

    file_exists = os.path.isfile(output_file)

    customer = ""

    extracted_data_in = extract_rows_in_rangeMed(input_file, "Container Movement - In", "Container Movement - Out")
    extracted_data_out = extract_rows_in_rangeMed(input_file, "Container Movement - Out", "Active Booking Listing Summary")

    with open(input_file, 'r') as infile, open(output_file, 'a', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        if file_exists:
            writer.writerow([])
            writer.writerow(['Report - ' + input])
            writer.writerow([])
        else:
            writer.writerow(['Report - ' + input])
            writer.writerow([])

        writer.writerow(['Gate In'])
        writer.writerow(['Customer', 'Date In', 'Container'])
        if len(extracted_data_in) > 1:
            for line in extracted_data_in:
                if len(line) > 2 and line[2] == '2EN8':
                    writer.writerow([line[4], line[5], line[1]])
        else:
            writer.writerow(['No tanks in.'])

        writer.writerow([])
        writer.writerow(['Gate Out'])
        writer.writerow(['Customer', 'Date Out', 'Container', 'Job No.'])
        if len(extracted_data_out) > 1:
            for line in extracted_data_out:
                if len(line) > 3 and line[3] == '2EN8':
                    writer.writerow([line[6], line[7], line[2], line[1]])
        else:
            writer.writerow(['No tanks out.'])

def ParseOutputsArc(input):
    input_file = input
    output_file = 'filtered_output.csv'

    file_exists = os.path.isfile(output_file)

    filedate = extract_date_from_filename(input_file)
    startdate = filedate - timedelta(days=3)

    extracted_data_in = extract_rows_in_range_xlsx(input_file, "INBOUND", "Disclaimer")
    extracted_data_out = extract_rows_in_range_xlsx(input_file, "TANKS IN DEPOT", "INBOUND")

    with open(input_file, 'r') as infile, open(output_file, 'a', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        if file_exists:
            writer.writerow([])
            writer.writerow(['Report - ' + input])
            writer.writerow([])
        else:
            writer.writerow(['Report - ' + input])
            writer.writerow([])

        writer.writerow(['Gate In'])
        writer.writerow(['Date In', 'Container'])
        if len(extracted_data_in) > 1:
            for line in extracted_data_in:
                if len(line) > 1 and line[1] != '' and line[1] != 'Status':
                    date_obj = datetime.strptime(line[4], "%Y-%m-%d %H:%M:%S")

                    writer.writerow([date_obj.strftime("%#d/%#m"), line[2]])
        else:
            writer.writerow(['No tanks in.'])

        writer.writerow([])
        writer.writerow(['Gate Out'])
        writer.writerow(['Date Out', 'Container'])
        if len(extracted_data_out) > 1:
            for line in extracted_data_out:
                if len(line) > 0 and line[1] != '' and line[9] != '':
                    if is_valid_date(line[9]):
                        linedate = datetime.strptime(line[9], "%Y-%m-%d %H:%M:%S")
                        
                        if startdate <= linedate <= filedate:
                            writer.writerow([linedate.strftime("%#d/%#m"), line[2]])

                   
        else:
            writer.writerow(['No tanks out.'])



def RunTSA():
    for filename in os.listdir("reports/TSA"):
        name, ext = os.path.splitext(filename)
        ParseOutputsTSA(ReadReport("reports/TSA/" + name))

def RunWTS():
    for filename in os.listdir("reports/WTS"):
        name, ext = os.path.splitext(filename)
        ParseOutputsWTS(ReadReport("reports/WTS/" + name))

def RunMED():
    for filename in os.listdir("reports/MED"):
        name, ext = os.path.splitext(filename)
        ParseOutputsMED(ReadReportMED("reports/MED/" + name))

def RunArc():
    for filename in os.listdir("reports/ARC"):
        name, ext = os.path.splitext(filename)
        ParseOutputsArc("reports/ARC/" + name + ".xlsx")

RunArc()
RunMED()
RunTSA()
RunWTS()
