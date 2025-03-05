from flask import Flask, render_template, request, send_file
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import pandas as pd
import re
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create uploads directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def sanitize_filename(filename):
    """Remove invalid characters and limit length."""
    filename = re.sub(r'[^A-Za-z0-9_\-]', '', filename)  # Removes invalid characters
    return filename[:255]  # Ensure filename does not exceed system limits

def extract_text_from_scanned_pdf(pdf_path):
    """Extract text from a scanned PDF using OCR."""
    text = ""
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        pix = page.get_pixmap()
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text += pytesseract.image_to_string(img) + "\n"
    
    return text

def extract_course_title(text):
    """Extract the course title from the OCR text, handling OCR errors."""
    lines = text.splitlines()
    for line in lines:
        if re.match(r'^[A-Z]{2,}-\d{3}-\d{2}', line):  # Matches course-like format (e.g., MIS-353-01)
            return re.sub(r'[^A-Za-z0-9\- ]', '', line.strip())  # Remove OCR artifacts
    return "Unknown Course"

def extract_assignments_and_grades(text):
    """Extract assignments, grades, and section weights from OCR text."""
    print("\n=== Starting Grade Extraction ===")
    print("Raw text first 500 characters:")
    print(text[:500])
    print("=== End of raw text sample ===\n")

    lines = text.splitlines()
    print(f"Number of lines: {len(lines)}")
    
    extracted_data = []
    
    print("\n=== Processing Lines ===")
    for i, line in enumerate(lines):
        line = line.strip()
        print(f"\nLine {i}: '{line}'")
        
        # Skip empty lines and header lines
        if not line or any(header in line for header in ["Grade tem", "Points", "Comments", "Grades -"]):
            print(f"Skipping line: {line}")
            continue

        # Try to extract grade information
        grade_pattern = re.search(r'(\d+)/(\d+)\s*(?:\d+%)?', line)
        if grade_pattern:
            achieved = float(grade_pattern.group(1))
            total = float(grade_pattern.group(2))
            
            # Get assignment name by removing the grade part
            assignment = line.replace(grade_pattern.group(0), '').strip()
            print(f"Found grade: Assignment='{assignment}', Achieved={achieved}, Total={total}")
            
            entry = {
                'Section': 'All Grades',
                'Item': assignment,
                'Achieved': achieved,
                'Total': total
            }
            extracted_data.append(entry)
            print(f"Added entry: {entry}")
    
    print("\n=== Extraction Summary ===")
    print(f"Total entries found: {len(extracted_data)}")
    
    if not extracted_data:
        print("WARNING: No grades were extracted!")
        return pd.DataFrame(columns=['Section', 'Item', 'Achieved', 'Total', 'Weight'])
    
    # Create DataFrame
    df = pd.DataFrame(extracted_data)
    df['Weight'] = 100.0  # Add default weight
    
    print("\n=== Final DataFrame ===")
    print(df)
    print("=== End of DataFrame ===\n")
    
    return df

def save_to_excel(grades_df, excel_path):
    """Save the grades DataFrame to an Excel file."""
    grades_df.to_excel(excel_path, index=False)

def calculate_overall_grade(grades_df):
    """Calculate the overall percentage grade using section weights."""
    if grades_df.empty:
        return 0
    
    # Calculate section grades
    section_grades = {}
    for section in grades_df['Section'].unique():
        section_data = grades_df[grades_df['Section'] == section]
        section_achieved = section_data['Achieved'].sum()
        section_total = section_data['Total'].sum()
        section_weight = section_data['Weight'].iloc[0] if not section_data['Weight'].isna().all() else None
        
        if section_total > 0:
            section_grades[section] = {
                'percentage': (section_achieved / section_total) * 100,
                'weight': section_weight
            }
    
    # Calculate weighted average if weights are available
    if any(grade['weight'] is not None for grade in section_grades.values()):
        weighted_sum = 0
        total_weight = 0
        
        for section, data in section_grades.items():
            if data['weight'] is not None:
                weighted_sum += data['percentage'] * (data['weight'] / 100)
                total_weight += data['weight']
        
        if total_weight > 0:
            return round(weighted_sum * (100 / total_weight), 2)
    
    # Fallback to simple average if no weights are found
    total_achieved = grades_df['Achieved'].sum()
    total_possible = grades_df['Total'].sum()
    return round((total_achieved / total_possible) * 100, 2) if total_possible > 0 else 0

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        print("\n=== Starting File Upload ===")
        
        # Check if a file was uploaded
        if 'file' not in request.files:
            print("No file part in request")
            return render_template('upload.html', error='No file uploaded')
        
        file = request.files['file']
        if file.filename == '':
            print("No selected file")
            return render_template('upload.html', error='No file selected')
        
        print(f"Processing file: {file.filename}")
        
        if file and allowed_file(file.filename):
            # Secure the filename and save the uploaded file
            filename = secure_filename(file.filename)
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            print(f"Saving file to: {pdf_path}")
            file.save(pdf_path)
            
            try:
                # Extract text from PDF
                print("Extracting text from PDF...")
                text = extract_text_from_scanned_pdf(pdf_path)
                print(f"Extracted text length: {len(text)}")
                
                # Extract course title
                print("Extracting course title...")
                course_title = extract_course_title(text)
                print(f"Found course title: {course_title}")
                
                # Extract grades
                print("Extracting grades...")
                grades_df = extract_assignments_and_grades(text)
                
                print("\n=== Grades DataFrame Info ===")
                print("DataFrame empty?", grades_df.empty)
                if not grades_df.empty:
                    print("Columns:", grades_df.columns.tolist())
                    print("Shape:", grades_df.shape)
                print("=== End of DataFrame Info ===\n")
                
                if grades_df.empty:
                    print("No grades found in PDF")
                    return render_template('upload.html', error='No grades found in the PDF')
                
                # Save to Excel
                excel_filename = f'grades_{sanitize_filename(filename[:-4])}.xlsx'
                excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
                print(f"Saving Excel file to: {excel_path}")
                save_to_excel(grades_df, excel_path)
                
                # Calculate overall grade
                print("Calculating overall grade...")
                overall_grade = calculate_overall_grade(grades_df)
                print(f"Overall grade: {overall_grade}")
                
                # Convert DataFrame to HTML table
                print("Converting DataFrame to HTML...")
                grades_table = grades_df.to_html(classes='table table-striped', index=False)
                
                return render_template('results.html',
                                     course_title=course_title,
                                     overall_grade=f"{overall_grade:.2f}%",
                                     grades_table=grades_table,
                                     excel_filename=excel_filename)
                
            except Exception as e:
                print(f"Error occurred: {str(e)}")
                import traceback
                print("Traceback:")
                print(traceback.format_exc())
                return render_template('upload.html', error=f'Error processing file: {str(e)}')
            
            finally:
                # Clean up the uploaded PDF
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                    print(f"Cleaned up PDF file: {pdf_path}")
        
        return render_template('upload.html', error='Invalid file type')
    
    return render_template('upload.html')

@app.route('/download/<filename>')
def download_file(filename):
    if not filename.endswith('.xlsx'):
        return "Invalid file type", 400
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(file_path):
        return f"Error: The file {filename} does not exist.", 404
    
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    from waitress import serve
    serve(app, host="0.0.0.0", port=8080)

