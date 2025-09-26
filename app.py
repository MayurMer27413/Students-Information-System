from flask import Flask, render_template, request, redirect, url_for, flash
import os
from werkzeug.utils import secure_filename
from upload_handler import handle_photo_upload
import openpyxl


app = Flask(__name__)
app.secret_key = 'change-this-to-a-random-secret'


EXCEL_FILE = 'students.xlsx'
ID_COLUMN = 'Enrollment Number' # column name in the excel which holds unique IDs
PHOTO_COLUMN = 'Profile Photo' # column name for student profile photos




def load_students():
    try:
        # Load the workbook and get the active sheet
        workbook = openpyxl.load_workbook(EXCEL_FILE)
        sheet = workbook.active
        
        # Get headers from the first row
        headers = [cell.value for cell in sheet[1]]
        
        # Create a dictionary to store all student data
        students_data = []
        
        # Iterate through rows (starting from row 2, as row 1 contains headers)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Create a dictionary for each student
            student = {}
            for i, value in enumerate(row):
                # Convert None to empty string
                student[headers[i]] = str(value) if value is not None else ''
            students_data.append(student)
            
        return students_data
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None





@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')




@app.route('/search', methods=['POST'])
def search():
    student_id = request.form.get('student_id', '').strip()
    if not student_id:
        flash('Please enter a student ID.')
        return redirect(url_for('index'))

    students = load_students()
    if students is None:
        flash('Database file not found. Make sure students.xlsx exists.')
        return redirect(url_for('index'))

    # Find student with matching ID (exact match)
    match = None
    for student in students:
        if student[ID_COLUMN].strip() == student_id:
            match = student
            break

    if match is None:
        flash('Student not found.')
        return redirect(url_for('index'))

    # Format DOB to remove time component if it exists
    if 'DOB' in match and '00:00:00' in match['DOB']:
        match['DOB'] = match['DOB'].split(' ')[0]

    # Redirect to student page passing data via POST-redirect is non-trivial, so render template directly
    return render_template('student.html', student=match)




# Add route to display the upload photo form
@app.route('/upload', methods=['GET'])
def upload():
    return render_template('upload_photo.html')


# Initialize the photo upload handler
handle_photo_upload(app)


if __name__ == '__main__':
    app.run(debug=True)