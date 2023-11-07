import os
from pyexcel_ods3 import get_data
from datetime import datetime
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.opendocument import load, OpenDocumentSpreadsheet

# Function to load responses from an ODS file
def load_responses_from_ods(file_path):
    responses = {}
    if os.path.exists(file_path):
        data = get_data(file_path)
        if data:
            sheet = data.get(list(data.keys())[0])  # Assuming responses are on the first sheet
            for row in sheet:
                if len(row) >= 2:
                    keyword = row[0].strip().lower()
                    response = row[1].strip()
                    responses[keyword] = response
    return responses

# Function to initialize an ODS document or load an existing one
def initialize_or_load_ods_file(file_path, roll_number):
    if os.path.exists(file_path):
        doc = load(file_path)
        table = doc.spreadsheet.getElementsByType(Table)[0]  # Get the first table (assuming it's there)
    else:
        doc = OpenDocumentSpreadsheet()
        table = Table(name="Conversation")
        doc.spreadsheet.addElement(table)

        # Add header row for roll number, date, time, user query, and assistant response
        header_row = TableRow()
        table.addElement(header_row)
        header_row.addElement(TableCell())
        header_row.addElement(TableCell())
        header_row.addElement(TableCell())
        header_row.addElement(TableCell())
        header_row.addElement(TableCell())

    return doc, table

# Function to log the conversation to an ODS file with the correct format
def log_conversation_to_ods(doc, table, roll_number, user_query, assistant_response):
    row = TableRow()
    table.addElement(row)

    roll_number_cell = TableCell()
    roll_number_cell.addElement(P(text=roll_number))
    row.addElement(roll_number_cell)

    timestamp_cell = TableCell()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    timestamp_cell.addElement(P(text=timestamp))
    row.addElement(timestamp_cell)

    user_cell = TableCell()
    user_cell.addElement(P(text=user_query))
    row.addElement(user_cell)

    assistant_cell = TableCell()
    assistant_cell.addElement(P(text=assistant_response))
    row.addElement(assistant_cell)

# Input the roll number from the user
roll_number = input("Please enter your roll number: ")

# Specify the path to your ODS files
ods_file_path_1 = 'repo/linux.ods'  # Linux commands
ods_file_path_2 = 'repo/smartVA.ods'  # lab_manual
ods_file_path_3 = 'repo/OUTput.ods'  # lab_solution

# Load responses from both ODS files
responses1 = load_responses_from_ods(ods_file_path_1)
responses2 = load_responses_from_ods(ods_file_path_2)
responses3 = load_responses_from_ods(ods_file_path_3)

# Initialize or load the ODS document and table
doc, table = initialize_or_load_ods_file("conversation.ods", roll_number)

# Function to process user queries
def process_query(query):
    query = query.lower()
    response = ""

    for keyword, resp in responses1.items():
        if keyword in query:
            response = resp
            break

    for keyword, resp in responses2.items():
        if keyword in query:
            response = resp
            break

    for keyword, resp in responses3.items():
        if keyword in query:
            response = resp
            break

    if not response:
        response = "I'm not sure how to help with that. Please ask another question."

    # Log the conversation to the existing ODS file
    log_conversation_to_ods(doc, table, roll_number, query, response)

    return response

# Main loop
while True:
    user_query = input("You: ")

    if user_query.lower() == "exit":
        break

    response = process_query(user_query)
    print("Assistant:", response)
    print("----")

# Save the updated conversation to the ODS file
doc.save("conversation.ods")

print("Goodbye!")
