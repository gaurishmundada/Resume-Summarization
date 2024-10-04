import os
from flask import Flask, render_template, request, send_file
import pandas as pd
import spacy
import PyPDF2
import re
import docx  # For handling DOCX files
from PyPDF2.errors import PdfReadError

# Load the Spacy model for entity recognition
nlp = spacy.load('en_core_web_sm')

# Initialize Flask app
app = Flask(__name__)

# Function to extract text from PDF file using its stream
def extract_text_from_pdf(pdf_file):
    try:
        reader = PyPDF2.PdfReader(pdf_file.stream)
        text = ''
        for page_num in range(len(reader.pages)):
            text += reader.pages[page_num].extract_text()
        return text
    except PdfReadError:
        return "Error reading the PDF. The file might be corrupt or incorrectly formatted."

# Function to extract text from DOCX file
def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

# Function to extract text from uploaded file (PDF or DOCX)
def extract_text_from_file(file):
    if file.filename.endswith('.pdf'):
        return extract_text_from_pdf(file)
    elif file.filename.endswith('.docx'):
        return extract_text_from_docx(file)
    else:
        return "Unsupported file type. Only PDF and DOCX are allowed."

# Function to extract the name using Spacy (first PERSON entity found)
def extract_name(text):
    doc = nlp(text)
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            return ent.text.strip()
    return "Name not found"

# Function to extract education, CGPA, and college name
def extract_education_and_cgpa(text):
    education_info = None
    college_name = None
    cgpa = None

    # Look for education details using broader patterns
    education_pattern = r'(Bachelor|Master).*?\n.*?(?=\n|$)'
    cgpa_pattern = r'CGPA[:\-]?\s*(\d+\.\d+)'
    college_pattern = r'(University|Institute|College).*?\n.*?(?=\n|$)'

    education_match = re.search(education_pattern, text, re.IGNORECASE)
    if education_match:
        education_info = education_match.group(0).strip()

    college_match = re.search(college_pattern, text, re.IGNORECASE)
    if college_match:
        college_name = college_match.group(0)

    cgpa_match = re.search(cgpa_pattern, text)
    if cgpa_match:
        cgpa = cgpa_match.group(1)

    return education_info, cgpa, college_name

# Function to extract skills
def extract_skills(text):
    skills_pattern = r'SKILLS\s*(.*?)\s*(EXPERIENCE|PROJECTS|CERTIFICATIONS|$)'
    skills_match = re.search(skills_pattern, text, re.DOTALL | re.IGNORECASE)
    if skills_match:
        skills = skills_match.group(1).replace("\n", ", ").strip()
        return skills
    return "Skills not found"

# Route for the home page
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle resume upload and extraction
@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        files = request.files.getlist('resumes')
        data = []

        for file in files:
            text = extract_text_from_file(file)

            name = extract_name(text)
            education, cgpa, college = extract_education_and_cgpa(text)
            skills = extract_skills(text)

            # Add data to the list
            data.append([name, college, education, cgpa, skills])

        # Convert the data to a DataFrame
        df = pd.DataFrame(data, columns=['Name', 'College Name', 'Education', 'CGPA', 'Skills'])
        output_file_csv = 'summarized_resumes.csv'
        output_file_excel = 'summarized_resumes.xlsx'

        # Save the data as both CSV and Excel
        df.to_csv(output_file_csv, index=False)
        df.to_excel(output_file_excel, index=False)

        # Render results in HTML table with a download option
        return render_template('results.html', tables=[df.to_html(classes='data', header="true")], csv_file=output_file_csv, excel_file=output_file_excel)

# Route to handle Excel file download
@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

# Run the Flask app
if __name__ == '__main__':
    app.run(debug=True)
