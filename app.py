from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import os
from werkzeug.utils import secure_filename
from upload_handler import handle_photo_upload


app = Flask(__name__)
app.secret_key = 'change-this-to-a-random-secret'


EXCEL_FILE = 'students.xlsx'
ID_COLUMN = 'Enrollment Number' # column name in the excel which holds unique IDs
PHOTO_COLUMN = 'Profile Photo' # column name for student profile photos




def load_students():
    try:
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl', dtype=str)
        df.fillna('', inplace=True)
        return df
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

    df = load_students()
    if df is None:
        flash('Database file not found. Make sure students.xlsx exists.')
        return redirect(url_for('index'))

    # Find row(s) with matching ID (exact match)
    match = df[df[ID_COLUMN].astype(str).str.strip() == student_id]

    if match.empty:
        flash('Student not found.')
        return redirect(url_for('index'))

    # Convert first matched row to a dict of text
    row = match.iloc[0].astype(str).to_dict()
    
    # Format DOB to remove time component if it exists
    if 'DOB' in row and '00:00:00' in row['DOB']:
        row['DOB'] = row['DOB'].split(' ')[0]

    # Redirect to student page passing data via POST-redirect is non-trivial, so render template directly
    return render_template('student.html', student=row)




# Add route to display the upload photo form
@app.route('/upload', methods=['GET'])
def upload():
    return render_template('upload_photo.html')


# Initialize the photo upload handler
handle_photo_upload(app)


if __name__ == '__main__':
    app.run(debug=True)