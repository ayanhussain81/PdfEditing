from flask import Flask, render_template, request
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PyPDF2.generic import NameObject
import PyPDF2

app = Flask(__name__)

# Google Sheets API setup
sheet_id = "1OySSaM59h61PQdcAtwGfZrnNQSGj73MP1aOZ1tg6YcQ"
sheet_name = "forms"
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

# Now you can use the 'url' variable to read the CSV into a Pandas DataFrame
df = pd.read_csv(url)

# PDF filling function (modify as needed)
def fill_pdf(input_pdf_path, output_pdf_path, field_data, target_page_index, checkbox_data):
    with open(input_pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        for page_index, page in enumerate(pdf_reader.pages):
            if page_index == target_page_index:
                # Check if the page has form fields
                if '/Annots' in page:
                    pdf_writer.update_page_form_field_values(page, field_data)
                    annotations = page['/Annots']
                    for annotation in annotations:
                        obj = annotation.get_object()
                        if '/FT' in obj and obj['/FT'] == '/Btn':  # Check if it's a button (checkbox)
                            field_name = obj.get('/T')
                            if field_name and field_name in checkbox_data:
                                field_value = checkbox_data[field_name]
                                if isinstance(field_value, bool):
                                    if field_value:
                                        obj.update({
                                            PyPDF2.generic.NameObject("/V"): PyPDF2.generic.NameObject("/Yes"),
                                            PyPDF2.generic.NameObject("/AS"): PyPDF2.generic.NameObject("/Yes")
                                        })
                                    else:
                                        obj.update({
                                            PyPDF2.generic.NameObject("/V"): PyPDF2.generic.NameObject("/Off"),
                                            PyPDF2.generic.NameObject("/AS"): PyPDF2.generic.NameObject("/Off")
                                        })

            pdf_writer.add_page(page)

        # Write filled PDF to the output file
        with open(output_pdf_path, 'wb') as output_file:
            pdf_writer.write(output_file)

# Flask routes
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        name = request.form.get("name")
        target_page_index = 1
            # Iterate through rows in the DataFrame
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
            # Create a new dictionary for the current row
            row_data = {}

            sheet_name = row['First Name'] +" "+ row['Last Name']
            if sheet_name.lower() == name.lower():  
                
                if row['Title'] in checkbox_title:
                    checkbox_data = {checkbox_title[row['Title']]: True}

                output_pdf_path = f"{row['First Name']}_{row['Last Name']}.pdf"  # Example output file name
                for column_index, column in enumerate(headers):
                    # Extract data from the corresponding column in the DataFrame using index
                    row_data[column] = row[column_index]
                fill_pdf("input.pdf", output_pdf_path, row_data, target_page_index, checkbox_data)
                message = f"PDF generated for {row['First Name']} {row['Last Name']}. Check the file: {output_pdf_path}"
                # return f"PDF generated for {row['First Name']} {row['Last Name']}. Check the file: {output_pdf_path}"
                return render_template("index.html", title="PDF Generator", message=message)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)