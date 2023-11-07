import os
import pandas as pd
from pyexcel_ods import save_data
import re


# Define a function to remove non-XML-compatible characters
def remove_non_xml_chars(text):
    return re.sub(r'[^\x09\x0A\x0D\x20-\x7E]', '', text)


# Define the directory to read files from
directory_path = '/home/zerry/PycharmProjects/VA/lab_response'

# Initialize data dictionary to store file names and content
data = {'File Name': [], 'Content': []}

# Iterate through the files in the directory
for filename in os.listdir(directory_path):
    if os.path.isfile(os.path.join(directory_path, filename)):
        try:
            with open(os.path.join(directory_path, filename), 'r', encoding='utf-8') as file:
                content = file.read()
        except UnicodeDecodeError:
            # Try opening the file with 'latin-1' encoding as a fallback
            with open(os.path.join(directory_path, filename), 'r', encoding='latin-1') as file:
                content = file.read()

        # Remove non-XML-compatible characters from the content
        cleaned_content = remove_non_xml_chars(content)

        data['File Name'].append(filename)
        data['Content'].append(cleaned_content)

# Create a DataFrame from the data dictionary
df = pd.DataFrame(data)

# Define the output ODS file path
ods_file_path = 'OUTput.ods'

# Save the DataFrame to an ODS file
save_data(ods_file_path, {"Sheet 1": df.to_dict(orient='split')['data']})

print(f'Data has been saved to {ods_file_path}')
