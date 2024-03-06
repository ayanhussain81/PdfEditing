from flask import Flask, render_template, request, send_file
import pandas as pd
import PyPDF2
from PyPDF2.generic import NameObject
import requests
from io import BytesIO
import tempfile
import os
from pathlib import Path




app = Flask(__name__)

# Get a temporary directory
temp_folder = Path(tempfile.gettempdir())

def read_google_sheet(sheet_id, sheet_name):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    return pd.read_csv(url)

def process_registration_data(value):
    result = []
    lines = value.strip().split('\n')
    headers = lines[:6]
    values = [line.split('\n') for line in lines[6:] if line.strip()]   # Exclude blank lines

    for i in range(0, len(values), len(headers)):
        chunk = values[i:i + len(headers)]
        row_dict = dict(zip(headers, [val[0] for val in chunk]))
        result.append(row_dict)

    return result

def process_result_data_page3(df,row_data,row, checkbox_data,result):
    count = 0

    for entry in result:
        type_value = entry.get('Type (registration/licence)', '')
        if type_value:
            key = f'Text Field 8{6 + count}'  
            row_data[key] = type_value  

        reg_value = entry.get('Regulator name (issuer of licence/registration)', '')
        if reg_value: 
            key = f'Text Field {89+ count}'  
            row_data[key] = reg_value  

        state_value = entry.get('State', '')    
        if state_value:
            key = f'Text Field 9{2+ count}'  
            row_data[key] = state_value  

        date_value = entry.get('Date (first issued)', '')
        if date_value:
            key = f'Text Field 9{5+ count}'  
            row_data[key] = date_value  

        number_value = entry.get('Number (registration/ licence)', '')
        if number_value:
            key = f'Text Field {98+ count}'  
            row_data[key] = number_value  

        lic_value = entry.get('Lic/Reg (certified)', '')
        if lic_value.lower() == 'yes':
            key = f'Check Box 13{1 + count}'
            checkbox_data[key] = True 
        count += 1
    row_data['Street no and name 6'] = row['registration number/s (if Yes)']

    checkbox_data['Check Box 120'] = row['Currently registered'].lower() == 'yes'
    checkbox_data['Check Box 119'] = not checkbox_data['Check Box 120']

    checkbox_data['Check Box 122'] = row['Are you currently authorised to perform building work outside of Victoria?'].lower() == 'yes'
    checkbox_data['Check Box 123'] = not checkbox_data['Check Box 122']

    checkbox_data['Check Box 124'] = row['Have you previously been (but not currently) authorised to perform building work outside of Victoria?'].lower() == 'yes'
    checkbox_data['Check Box 125'] = not checkbox_data['Check Box 124']

    checkbox_data['Check Box 126'] = row['Do you hold a current licence to perform high risk work issued by an Australian state or territory workplace health and safety regulator?'].lower() == 'yes'
    checkbox_data['Check Box 127'] = not checkbox_data['Check Box 126']

    checkbox_data['Check Box 128'] = row['Do you hold a current Construction Induction Card (White Card) issued by an Australian state or territory workplace health and safety regulator?'].lower() == 'yes'
    checkbox_data['Check Box 129'] = not checkbox_data['Check Box 128']

    return row_data, checkbox_data


def process_qualification_data(value):
    result = []
    lines = value.strip().split('\n')
    headers = lines[:5]
    values = [line.split('\n') for line in lines[5:] if line.strip()]   # Exclude blank lines

    for i in range(0, len(values), len(headers)):
        chunk = values[i:i + len(headers)]
        row_dict = dict(zip(headers, [val[0] for val in chunk]))
        result.append(row_dict)

    return result

def process_result_data_page4(row_data, checkbox_data,result):
    count = 0

    for entry in result:
        qualification_value = entry.get('Name of qualification', '')
        if qualification_value:
            key = f'Text Field {116 + count}'  
            row_data[key] = qualification_value  

        institution_value = entry.get('Institution (TAFE, University, RTO)', '')
        if institution_value: 
            key = f'Text Field {124+ count}'  
            row_data[key] = institution_value  

        year_value = entry.get('Year completed', '')    
        if year_value:
            key = f'Text Field {129+ count}'  
            row_data[key] = year_value  

        check_qualification_value = entry.get('Qualification', '')
        if check_qualification_value.lower() == 'yes':
            key = f'Check Box {135 + count}'
            checkbox_data[key] = True 
            
        check_result_value = entry.get('Results', '')
        if check_result_value.lower() == 'yes':
            key = f'Check Box {141 + count}'
            checkbox_data[key] = True 

        count += 1
    return row_data, checkbox_data


def process_result_data_page2(df_page2, row_data,row, checkbox_data):
    headers = ["Title", "First name", "Middle name", "Surname", "Date of birth", "Email", "Mobile number",
               "Home telephone number", "Street no and name", 'Suburb', 'State', 'Postcode',
               'Street no and name 3', 'Suburb 3', 'State 4', 'Postcode 4',
               'Street no and name 4', 'Suburb 4', 'State 5', 'Postcode 5',
               'Street no and name 5', 'Suburb 5']

    # Create a new dictionary for the current row
    checkbox_title = {
        "Mr": "Check Box 3",
        "Mrs": "Check Box 4",
        "Ms": "Check Box 5",
        "Miss": "Check Box 6",
    }

    if row['Title'] in checkbox_title:
        checkbox_data = {checkbox_title[row['Title']]: True}

    count = 0
    for column in headers:
        try:
            # Extract data from the corresponding column in the DataFrame using column name
            row_data[column] = df_page2.iloc[0, count]
        except KeyError as e:
            print(f"KeyError: {e} not found in df_page2")
        count += 1

    return row_data, checkbox_data


def download_pdf(url):
  response = requests.get(url)
  if response.status_code == 200:
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
      temp_file.write(response.content)
      return temp_file.name
  else:
    raise Exception(f"Failed to download PDF from {url}")

# PDF filling function
def fill_pdf(input_pdf_path, output_pdf_path, field_data, checkbox_data):
    with open(input_pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        for page_index, page in enumerate(pdf_reader.pages):
            if '/Annots' in page:
                pdf_writer.update_page_form_field_values(page, field_data)
                annotations = page['/Annots']
                for annotation in annotations:
                    obj = annotation.get_object()
                    if '/FT' in obj and obj['/FT'] == '/Btn':
                        field_name = obj.get('/T')
                        if field_name and field_name in checkbox_data:
                            field_value = checkbox_data[field_name]
                            if isinstance(field_value, bool):
                                if field_value:
                                    obj.update({
                                        NameObject("/V"): NameObject("/Yes"),
                                        NameObject("/AS"): NameObject("/Yes")
                                    })
                                else:
                                    obj.update({
                                        NameObject("/V"): NameObject("/Off"),
                                        NameObject("/AS"): NameObject("/Off")
                                    })

            pdf_writer.add_page(page)

        with open(output_pdf_path, 'wb') as output_file:
            pdf_writer.write(output_file)

# Flask routes
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        sheet_id = "1OySSaM59h61PQdcAtwGfZrnNQSGj73MP1aOZ1tg6YcQ"
        sheet_name = "forms"

        name = request.form.get("name")
        df = read_google_sheet(sheet_id, sheet_name)

        for index, row in df.iterrows():
            sheet_name = row['First Name'] + " " + row['Last Name']
            if sheet_name.lower() == name.lower():  
                row_data, checkbox_data={},{}
                df_page2 = df.iloc[[index], :22]

                # For page 2
                row_data, checkbox_data = process_result_data_page2(df_page2,row_data,row, checkbox_data)
            
                # For page 3
                registration_values = row['Registration (Repeater)']
                try:
                    result_data_3 = process_registration_data(registration_values)
                    row_data, checkbox_data = process_result_data_page3(df,row_data,row, checkbox_data,result_data_3)
                except:
                    pass
                # For page 4
                try:
                    qualifications_values = row['Qualification Field (Repeater)']
                    result_data = process_qualification_data(qualifications_values)
                    row_data, checkbox_data = process_result_data_page4(row_data, checkbox_data,result_data)
                except:
                    pass


                output_pdf_path = f"{row['First Name']}_{row['Last Name']}.pdf"

                input_pdf_url = "https://github.com/ayanhussain81/PdfEditing/raw/main/input.pdf"
                input_pdf_content = download_pdf(input_pdf_url)

                filled_pdf_path = temp_folder / output_pdf_path

                fill_pdf(input_pdf_content, filled_pdf_path, row_data, checkbox_data)
                return send_file(filled_pdf_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)