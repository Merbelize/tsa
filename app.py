from flask import Flask, request, render_template, send_from_directory
from docx import Document
import os

app = Flask(__name__)

@app.route('/')
def form():
    return render_template('/template/index.html')

@app.route('/submit', methods=['POST'])
def submit_form():
    # Assuming your form has fields 'naam', 'datum', 'opdrachtgever', 'locatie', 'contactpersoon',
    # 'melding', 'oplossing', 'opvolging'
    naam = request.form['naam']
    datum = request.form['datum']  # Make sure to include fields like this in your form
    opdrachtgever = request.form['opdrachtgever']
    locatie = request.form['locatie']
    contactpersoon = request.form['contactpersoon']
    melding = request.form['melding']
    oplossing = request.form['oplossing']
    opvolging = request.form['opvolging']

    # Load the template
    doc = Document('documents/dagrapport_veiligheid_template.docx')
    
    # Replace placeholders
    for paragraph in doc.paragraphs:
        if '{naam_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{naam_tekst}', naam)
        if '{datum_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{datum_tekst}', datum)
        if '{opdrachtgever_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{opdrachtgever_tekst}', opdrachtgever)
        if '{locatie_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{locatie_tekst}', locatie)
        if '{contactpersoon_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{contactpersoon_tekst}', contactpersoon)
        if '{melding_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{melding_tekst}', melding)
        if '{oplossing_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{oplossing_tekst}', oplossing)
        if '{opvolging_tekst}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{opvolging_tekst}', opvolging)
        # Continue this process for the other placeholders

    # Save the document with a new name, e.g., based on a timestamp or an incrementing identifier
    file_path = '/submitted/' + 'dagrapport_' + naam + '.docx'
    doc.save(file_path)
    
    # Optionally, return the file directly to the user or a confirmation message
    return send_from_directory(directory=os.path.dirname(file_path), filename=os.path.basename(file_path), as_attachment=True)
    # or return "Form Submitted and Document Generated!"

if __name__ == '__main__':
    app.run(debug=True)
