from flask import Flask, render_template, request, send_file
import pandas as pd
import PyPDF2
from PyPDF2.generic import NameObject
import requests
from io import BytesIO
import tempfile



app = Flask(__name__)

# Google Sheets API setup
sheet_id = "1OySSaM59h61PQdcAtwGfZrnNQSGj73MP1aOZ1tg6YcQ"
sheet_name = "forms"
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

# Now you can use the 'url' variable to read the CSV into a Pandas DataFrame
df = pd.read_csv(url)

def download_pdf(url):
  response = requests.get(url)
  if response.status_code == 200:
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
      temp_file.write(response.content)
      return temp_file.name
  else:
    raise Exception(f"Failed to download PDF from {url}")

# PDF filling function
def fill_pdf(input_pdf_path, output_pdf_path, field_data, target_page_index, checkbox_data):
    with open(input_pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        for page_index, page in enumerate(pdf_reader.pages):
            if page_index == target_page_index:
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
        name = request.form.get("name")
        target_page_index = 1
        headers = ["Title", "First name", "Middle name", "Surname", "Date of birth", "Email", "Mobile number", 
                   "Home telephone number", "Street no and name", 'Suburb', 'State', 'Postcode', 
                   'Street no and name 3', 'Suburb 3', 'State 4', 'Postcode 4', 
                   'Street no and name 4', 'Suburb 4', 'State 5', 'Postcode 5', 
                   'Street no and name 5', 'Suburb 5']
        checkbox_title = {
                "Mr": "Check Box 3",
                "Mrs": "Check Box 4",
                "Ms": "Check Box 5",
                "Miss": "Check Box 6",
            }

        for index, row in df.iterrows():
            row_data = {}
            sheet_name = row['First Name'] + " " + row['Last Name']
            if sheet_name.lower() == name.lower():  
                if row['Title'] in checkbox_title:
                    checkbox_data = {checkbox_title[row['Title']]: True}

                output_pdf_path = f"{row['First Name']}_{row['Last Name']}.pdf"
                for column_index, column in enumerate(headers):
                    row_data[column] = row[column_index]
                input_pdf_url = "https://github.com/ayanhussain81/PdfEditing/raw/main/input.pdf"
                input_pdf_content = download_pdf(input_pdf_url)

                fill_pdf(input_pdf_content, output_pdf_path, row_data, target_page_index, checkbox_data)
                return send_file(output_pdf_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)