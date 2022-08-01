from flask import Flask, request
from docx import Document
import os
import spacy
from flask_cors import CORS
from werkzeug.utils import secure_filename
from spacy.matcher import Matcher

# Document upload configurations
UPLOAD_FOLDER = 'C:/Users/Sadunika/Desktop/Research/Project-Demo'
app = Flask(__name__)
CORS(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# predefined data
field_name_column_name = 'field name'
field_type_column_name = 'field type'
ignore_labels = ['fields', 'buttons', 'field name']

nlp = spacy.load("en_core_web_sm")
lemmatizer = nlp.get_pipe("lemmatizer")

# Rule-based entity recognition for UI attribute
ruler = nlp.add_pipe("entity_ruler")
# supported UI elements
patterns = [{"label": "UI_TEXT_FIELD", "pattern": "text field"},
            {"label": "UI_RADIO", "pattern": "radio"},
            {"label": "UI_DROPDOWN", "pattern": "dropdown"},
            {"label": "UI_CHECKBOX", "pattern": "checkbox"},
            {"label": "UI_DATE", "pattern": "date"},
            {"label": "UI_BUTTON", "pattern": "button"}]
ruler.add_patterns(patterns)

# Token-based matching for mandotory field recognition
matcher = Matcher(nlp.vocab)
pattern = [{"LOWER": "mandatory"}]
matcher.add("MandotoryPattern", [pattern])


@app.route('/upload', methods=['POST'])
def fileUpload():
    # access the uploaded file
    target = os.path.join(UPLOAD_FOLDER, 'test_docs')
    file = request.files['file']
    filename = secure_filename(file.filename)
    destination = "/".join([target, filename])
    file.save(destination)
    # access the extracted UI elements array
    ui_elements_array = extract_data(destination)
    response = {
        "data": ui_elements_array
    }
    # return the responce that consists of UI elements
    return response


def extract_data(destination):
    # access the uploaded document
    document = Document(destination)
    tables = []
    data = []
    field_name_column = None
    field_type_column = None

    # go throght the filed specification table of the document and add columns and rows to the data frame
    for table in document.tables:
        df = [['' for i in range(len(table.columns))]
              for j in range(len(table.rows))]
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                if cell.text:
                    df[i][j] = cell.text
    tables.append(df)

    # get field name column index
    field_name_column = get_column(table, field_name_column_name)
    # get field type column index
    field_type_column = get_column(table, field_type_column_name)
    # check if the field type index and filed name index available
    if field_name_column is not None and field_type_column is not None:
        # go throuth the table rows and extract UI element type, label name of the fields and if the field is mandatory or not
        for cell in tables:
            for d in cell:
                item = {
                    "element": extract_attribute(d[field_type_column]),
                    "label": extract_label(d[field_name_column]),
                    "mandotory": extract_mandatory_fields(d[field_type_column]),
                }
                if None not in ([item['label'], item['element']]):
                    data.append(item)
    # return the extracted UI elements list
        response = {
            "elements": data,
            "error": None
        }
        return response
    # if field type or filed name columns do not exist in the table, return an error message
    else:
        response = {
            "elements": [],
            "error": "Invalid columns"
        }
        return response


def extract_attribute(field_description):
    button_element = 'UI_BUTTON'
    if not field_description:
        return button_element
    else:
        # preparing data for tokenizing
        description = nlp(" ".join(field_description.lower().split()))
        # lemmatization process
        description_after_lemmatizer = nlp(
            ' '.join(map(str, [token.lemma_ for token in description])))
        # tokenization , remove stop words and unnecessary characters
        tokens = [
            token.text for token in description_after_lemmatizer if not token.is_punct if not token.is_stop]
        # preparation for entity recognition
        description_after_text_processing = nlp(' '.join(map(str, tokens)))
        # entity recognition
        attribute = ' '.join(map(
            str, [ent.label_ for ent in description_after_text_processing.ents if ent.label_ != 'CARDINAL']))
        return attribute


def extract_mandatory_fields(field_description):
    # preparing data for token matcher
    doc = nlp(field_description)
    matches = matcher(doc)
    # matching the token and return true or false
    for match_id, start, end in matches:
        if nlp.vocab.strings[match_id]:
            return True
        else:
            return False


def get_column(table, column_name):
    # return column index of the inserted column name
    for index, cell in enumerate(table.rows[0].cells):
        if cell.text.lower() == column_name:
            return index


def extract_label(field_description):
    # return label of the UI element
    field_data = field_description.lower()
    if field_data not in ignore_labels:
        return field_data.title()


if __name__ == "__main__":
    app.run()
