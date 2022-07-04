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

# Rule-based entity recognition for UI attribute
nlp = spacy.load("en_core_web_sm")
lemmatizer = nlp.get_pipe("lemmatizer")
ruler = nlp.add_pipe("entity_ruler")
# supported UI elements
patterns = [{"label": "UI_TEXT_FIELD", "pattern": "text field"}, {"label": "UI_RADIO", "pattern": "radio"},
            {"label": "UI_DROPDOWN", "pattern": "dropdown"}, {
                "label": "UI_CHECKBOX", "pattern": "checkbox"},
            {"label": "UI_DATE", "pattern": "date"}, {"label": "UI_BUTTON", "pattern": "button"}]
ruler.add_patterns(patterns)

# Token-based matching for mandotory field recognition
matcher = Matcher(nlp.vocab)
pattern = [{"LOWER": "mandatory"}]
matcher.add("MandotoryPattern", [pattern])


@app.route('/upload', methods=['POST'])
def fileUpload():
    target = os.path.join(UPLOAD_FOLDER, 'test_docs')
    file = request.files['file']
    filename = secure_filename(file.filename)
    destination = "/".join([target, filename])
    file.save(destination)
    array = extract_data(destination)
    response = {
        "data": array
    }
    return response


def extract_data(destination):
    document = Document(destination)
    tables = []
    data = []
    field_name_column = None
    field_type_column = None
    for table in document.tables:
        df = [['' for i in range(len(table.columns))]
              for j in range(len(table.rows))]
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                if cell.text:
                    df[i][j] = cell.text
    tables.append(df)
    field_name_column = get_column(table, field_name_column_name)
    field_type_column = get_column(table, field_type_column_name)
    if field_name_column is not None and field_type_column is not None:
        for cell in tables:
            for d in cell:
                item = {
                    "element": extract_attribute(d[field_type_column]),
                    "label": extract_label(d[field_name_column]),
                    "mandotory": extract_mandotory_fields(d[field_type_column]),
                }
                if None not in ([item['label'], item['element']]):
                    data.append(item)
        response = {
            "elements": data,
            "error": None
        }
        return response
    else:
        response = {
            "elements": [],
            "error": "Field Name and Field Type columns should be included in the Field Specification table"
        }
        return response


def extract_attribute(field_description):
    button_element = 'UI_BUTTON'
    if not field_description:
        return button_element
    else:
        description = nlp(" ".join(field_description.lower().split()))
        description_after_lemmatizer = nlp(
            ' '.join(map(str, [token.lemma_ for token in description])))
        tokens = [
            token.text for token in description_after_lemmatizer if not token.is_punct if not token.is_stop]
        description_after_text_processing = nlp(' '.join(map(str, tokens)))
        attribute = ' '.join(map(
            str, [ent.label_ for ent in description_after_text_processing.ents if ent.label_ != 'CARDINAL']))
        return attribute


def extract_mandotory_fields(field_description):
    doc = nlp(field_description)
    matches = matcher(doc)
    for match_id, start, end in matches:
        if nlp.vocab.strings[match_id]:
            return True
        else:
            return False


def get_column(table, column_name):
    for index, cell in enumerate(table.rows[0].cells):
        if cell.text.lower() == column_name:
            return index


def extract_label(field_description):
    field_data = field_description.lower()
    if field_data not in ignore_labels:
        return field_data.title()


if __name__ == "__main__":
    app.run()
