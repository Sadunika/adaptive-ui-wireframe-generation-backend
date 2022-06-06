from flask import Flask, request
from docx import Document
import os
from flask_cors import CORS
from numpy import equal
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'C:/Users/Sadunika/Desktop/Research/Project-Demo'
app = Flask(__name__)
CORS(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# supported attributes
supported_attribute_list = ['text field', 'radio', 'dropdown','checkbox','date']

def extract_data(destination):
    document = Document(destination)
    tables = []
    data = []
    for table in document.tables:
        df = [['' for i in range(len(table.columns))] for j in range(len(table.rows))]
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                if cell.text:
                    df[i][j] = cell.text
    tables.append(df)    
    for cell in tables:
        for d in cell:
            item =  {
                "element": extract_attribute(supported_attribute_list, d[2]),
                "label": extract_label(d[0]),
                "mandotory": extract_mandotory_fields(d[2]),
            }
            if None not in ([item['label'],item['element']]):
                data.append(item) 
    return data
    
@app.route('/upload', methods=['POST'])
def fileUpload():
    target=os.path.join(UPLOAD_FOLDER,'test_docs')
    file = request.files['file'] 
    filename = secure_filename(file.filename)
    destination="/".join([target, filename])
    file.save(destination)
    array = extract_data(destination)
    response = {
            "data": array
        }
    return response

def extract_attribute(supported_elements_list, field_description):
    button_element = 'button'
    if not field_description :
        return button_element
    else : 
        field_data = field_description.lower()
        for element_type in supported_elements_list:
            if element_type in field_data:
                return element_type
            

def extract_mandotory_fields(field_description):
    field_data = field_description.lower()
    if "mandatory" in field_data:
        return True
    else :
        return False

def extract_label(field_description):
    field_data = field_description.lower()
    ignore_labels = ['fields','buttons','field name']
    if field_data not in ignore_labels: return field_data.title()

if __name__ == "__main__":
    app.run()





