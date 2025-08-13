import os
import logging
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
from seating_algorithm import SeatingAlgorithm
from excel_handler import ExcelHandler
import json
from datetime import datetime
from database import (
    get_db, init_db, close_db, get_all_students, get_student_by_roll, add_student, 
    update_student, delete_student, get_all_rooms, get_room_by_id, add_room, 
    update_room, delete_room, create_seating_plan, add_seating_assignment, 
    add_unassigned_student, get_seating_plan, get_latest_seating_plan, 
    get_room_assignments, get_all_seating_plans, delete_seating_plan
)

# Set up logging
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key-change-in-production")

# Configuration
UPLOAD_FOLDER = 'uploads'
EXPORT_FOLDER = 'exports'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXPORT_FOLDER'] = EXPORT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)

# Initialize database
with app.app_context():
    init_db()

# Initialize handlers
excel_handler = ExcelHandler()
seating_algorithm = SeatingAlgorithm()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.teardown_appcontext
def teardown_db(exception):
    close_db()

@app.route('/')
def index():
    student_count = len(get_all_students())
    room_count = len(get_all_rooms())
    plans = get_all_seating_plans()
    return render_template('index.html', student_count=student_count, room_count=room_count, plans=plans)

@app.route('/upload_students', methods=['POST'])
def upload_students():
    try:
        if 'file' not in request.files:
            flash('No file selected', 'error')
            return redirect(url_for('index'))
        
        file = request.files['file']
        if file.filename == '':
            flash('No file selected', 'error')
            return redirect(url_for('index'))
        
        if file and file.filename and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Process the Excel file
            students_df = excel_handler.read_student_data(filepath)
            if students_df is not None:
                # Add students to database
                added = 0
                duplicates = 0
                for _, row in students_df.iterrows():
                    success = add_student(
                        str(row['roll_number']),
                        row['name'],
                        row['class'],
                        row['section'],
                        row['gender'].lower()
                    )
                    if success:
                        added += 1
                    else:
                        duplicates += 1
                
                flash(f'Successfully uploaded {added} students ({duplicates} duplicates skipped)', 'success')
            else:
                flash('Error processing Excel file. Please check the format.', 'error')
        else:
            flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)', 'error')
            
    except Exception as e:
        logging.error(f"Error uploading file: {str(e)}")
        flash(f'Error uploading file: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/add_manual_student', methods=['POST'])
def add_manual_student():
    try:
        roll_number = request.form.get('roll_number', '').strip()
        name = request.form.get('name', '').strip()
        class_name = request.form.get('class_name', '').strip()
        section = request.form.get('section', '').strip()
        gender = request.form.get('gender', '').strip().lower()
        
        if not all([roll_number, name, class_name, section, gender]):
            flash('All fields are required', 'error')
            return redirect(url_for('index'))
        
        # Add new student
        success = add_student(roll_number, name, class_name, section, gender)
        if success:
            flash('Student added successfully', 'success')
        else:
            flash('Roll number already exists', 'error')
        
    except Exception as e:
        logging.error(f"Error adding manual student: {str(e)}")
        flash(f'Error adding student: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/add_room', methods=['POST'])
def add_room():
    try:
        room_id = request.form.get('room_id', '').strip()
        rows = int(request.form.get('rows', 0))
        columns = int(request.form.get('columns', 0))
        extra_desks = int(request.form.get('extra_desks', 0))
        
        if not room_id or rows <= 0 or columns <= 0:
            flash('Invalid room configuration', 'error')
            return redirect(url_for('index'))
        
        # Add new room
        success = add_room(room_id, rows, columns, extra_desks)
        if success:
            base_capacity = rows * columns * 2
            total_capacity = base_capacity + (extra_desks * 2)
            flash(f'Room added successfully with {rows}×{columns} grid + {extra_desks} extra desks (Total: {total_capacity} students)', 'success')
        else:
            flash('Room ID already exists', 'error')
        
    except ValueError:
        flash('Invalid number format for rows, columns, or extra desks', 'error')
    except Exception as e:
        logging.error(f"Error adding room: {str(e)}")
        flash(f'Error adding room: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/generate_seating', methods=['POST'])
def generate_seating():
    try:
        # Get configuration options
        selected_classes = request.form.getlist('selected_classes')
        boy_girl_pairing = 'boy_girl_pairing' in request.form
        gender_separation = 'gender_separation' in request.form
        random_assignment = 'random_assignment' in request.form
        
        # Get students from database
        students = get_all_students()
        if not students:
            flash('No student data available. Please upload or add students first.', 'error')
            return redirect(url_for('index'))
        
        # Filter students if specific classes selected
        if selected_classes:
            students = [s for s in students if s['class'] in selected_classes]
        
        if not students:
            flash('No students match the selected criteria.', 'error')
            return redirect(url_for('index'))
        
        # Get rooms from database
        rooms = get_all_rooms()
        if not rooms:
            flash('No rooms configured. Please add rooms first.', 'error')
            return redirect(url_for('index'))
        
        # Convert to format expected by seating algorithm
        students_list = [dict(s) for s in students]
        rooms_list = [dict(r) for r in rooms]
        
        # Generate seating arrangement
        seating_plan, unassigned = seating_algorithm.generate_seating_plan(
            students_list, rooms_list, boy_girl_pairing, gender_separation, random_assignment
        )
        
        if not seating_plan:
            flash('Could not generate seating plan. Please check room capacity and constraints.', 'error')
            return redirect(url_for('index'))
        
        # Store results in database
        config = {
            'selected_classes': selected_classes,
            'boy_girl_pairing': boy_girl_pairing,
            'gender_separation': gender_separation,
            'random_assignment': random_assignment
        }
        
        plan_name = f"Seating Plan {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        plan_id = create_seating_plan(plan_name, json.dumps(config))
        
        # Store assignments
        for room_id, desks in seating_plan.items():
            for desk_number, student in desks.items():
                if student:  # Skip empty desks
                    add_seating_assignment(plan_id, room_id, desk_number, student['id'])
        
        # Store unassigned students
        for student in unassigned:
            add_unassigned_student(plan_id, student['id'])
        
        flash(f'Seating plan generated successfully! {len(unassigned)} students remain unassigned.', 'success')
        return redirect(url_for('preview', plan_id=plan_id))
        
    except Exception as e:
        logging.error(f"Error generating seating: {str(e)}")
        flash(f'Error generating seating plan: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/preview/<int:plan_id>')
def preview(plan_id):
    seating_data = get_seating_plan(plan_id)
    if not seating_data:
        flash('Seating plan not found', 'error')
        return redirect(url_for('index'))
    
    rooms = get_all_rooms()
    
    # Organize assignments by room
    room_assignments = {}
    for room in rooms:
        assignments = get_room_assignments(plan_id, room['room_id'])
        room_assignments[room['room_id']] = assignments
    
    return render_template('preview.html', 
                         plan_id=plan_id,
                         plan_info=seating_data['plan_info'],
                         room_assignments=room_assignments,
                         unassigned=seating_data['unassigned'], 
                         rooms=rooms)

@app.route('/export_room/<int:plan_id>/<room_id>')
def export_room(plan_id, room_id):
    try:
        seating_data = get_seating_plan(plan_id)
        if not seating_data:
            flash('Seating plan not found', 'error')
            return redirect(url_for('index'))
        
        room = get_room_by_id(room_id)
        if not room:
            flash('Room not found', 'error')
            return redirect(url_for('preview', plan_id=plan_id))
        
        # Get assignments for this room
        assignments = get_room_assignments(plan_id, room_id)
        
        # Convert to format expected by ExcelHandler
        room_seating = {}
        for assignment in assignments:
            room_seating[assignment['desk_number']] = dict(assignment)
        
        # Generate Excel file
        filename = excel_handler.export_room_seating(room_id, dict(room), room_seating)
        if filename:
            return send_file(filename, as_attachment=True)
        else:
            flash('Error generating Excel file', 'error')
            return redirect(url_for('preview', plan_id=plan_id))
            
    except Exception as e:
        logging.error(f"Error exporting room: {str(e)}")
        flash(f'Error exporting room: {str(e)}', 'error')
        return redirect(url_for('preview', plan_id=plan_id))

@app.route('/export_room_grid/<int:plan_id>/<room_id>')
def export_room_grid(plan_id, room_id):
    try:
        seating_data = get_seating_plan(plan_id)
        if not seating_data:
            flash('Seating plan not found', 'error')
            return redirect(url_for('index'))
        
        room = get_room_by_id(room_id)
        if not room:
            flash('Room not found', 'error')
            return redirect(url_for('preview', plan_id=plan_id))
        
        # Get assignments for this room
        assignments = get_room_assignments(plan_id, room_id)
        
        # Convert to format expected by ExcelHandler
        room_seating = {}
        for assignment in assignments:
            room_seating[assignment['desk_number']] = dict(assignment)
        
        # Generate Excel file with grid layout
        filename = excel_handler.export_room_grid_layout(room_id, dict(room), room_seating)
        if filename:
            return send_file(filename, as_attachment=True)
        else:
            flash('Error generating grid layout file', 'error')
            return redirect(url_for('preview', plan_id=plan_id))
            
    except Exception as e:
        logging.error(f"Error exporting grid layout: {str(e)}")
        flash(f'Error exporting grid layout: {str(e)}', 'error')
        return redirect(url_for('preview', plan_id=plan_id))

@app.route('/export_all_plans_zip/<int:plan_id>')
def export_all_plans_zip(plan_id):
    try:
        seating_data = get_seating_plan(plan_id)
        if not seating_data:
            flash('Seating plan not found', 'error')
            return redirect(url_for('index'))
        
        rooms = get_all_rooms()
        if not rooms:
            flash('No rooms found', 'error')
            return redirect(url_for('preview', plan_id=plan_id))
        
        # Prepare all room data
        all_room_data = []
        for room in rooms:
            assignments = get_room_assignments(plan_id, room['room_id'])
            room_seating = {}
            for assignment in assignments:
                room_seating[assignment['desk_number']] = dict(assignment)
            all_room_data.append({
                'room_id': room['room_id'],
                'room_info': dict(room),
                'assignments': room_seating
            })
        
        # Generate ZIP file
        zip_filename = excel_handler.export_all_rooms_zip(plan_id, all_room_data)
        if zip_filename:
            return send_file(zip_filename, as_attachment=True)
        else:
            flash('Error generating ZIP file', 'error')
            return redirect(url_for('preview', plan_id=plan_id))
            
    except Exception as e:
        logging.error(f"Error exporting all plans: {str(e)}")
        flash(f'Error exporting all plans: {str(e)}', 'error')
        return redirect(url_for('preview', plan_id=plan_id))

@app.route('/get_student_data')
def get_student_data():
    students = get_all_students()
    return jsonify([dict(s) for s in students])

@app.route('/get_rooms_data')
def get_rooms_data():
    rooms = get_all_rooms()
    return jsonify([dict(r) for r in rooms])

@app.route('/load_sample_data', methods=['POST'])
def load_sample_data():
    try:
        # Load sample student data
        if os.path.exists('jnv_students_test_data.xlsx'):
            df = pd.read_excel('jnv_students_test_data.xlsx')
            
            # Add sample students
            added = 0
            duplicates = 0
            for _, row in df.iterrows():
                success = add_student(
                    str(row['roll_number']),
                    row['name'],
                    row['class'],
                    row['section'],
                    row['gender'].lower()
                )
                if success:
                    added += 1
                else:
                    duplicates += 1
            
            # Add sample rooms
            sample_rooms = [
                ('EXAM-HALL-A', 6, 8, 2),
                ('EXAM-HALL-B', 5, 7, 3),
                ('LIBRARY-HALL', 8, 6, 0),
                ('COMPUTER-LAB', 4, 8, 2),
            ]
            
            for room in sample_rooms:
                add_room(*room)
            
            flash(f'Sample data loaded: {added} students (skipped {duplicates} duplicates) and 4 exam halls configured', 'success')
        else:
            flash('Sample data file not found. Please upload student data manually.', 'error')
    except Exception as e:
        logging.error(f"Error loading sample data: {str(e)}")
        flash('Error loading sample data', 'error')
    
    return redirect(url_for('index'))

@app.route('/clear_data', methods=['POST'])
def clear_data():
    data_type = request.form.get('data_type')
    
    if data_type == 'students':
        db = get_db()
        db.execute("DELETE FROM students")
        db.commit()
        flash('Student data cleared', 'info')
    elif data_type == 'rooms':
        db = get_db()
        db.execute("DELETE FROM rooms")
        db.commit()
        flash('Room data cleared', 'info')
    elif data_type == 'seating':
        db = get_db()
        db.execute("DELETE FROM seating_plans")
        db.execute("DELETE FROM seating_assignments")
        db.execute("DELETE FROM unassigned_students")
        db.commit()
        flash('All seating plans cleared', 'info')
    elif data_type == 'all':
        db = get_db()
        db.execute("DELETE FROM students")
        db.execute("DELETE FROM rooms")
        db.execute("DELETE FROM seating_plans")
        db.execute("DELETE FROM seating_assignments")
        db.execute("DELETE FROM unassigned_students")
        db.commit()
        flash('All data cleared', 'info')
    
    return redirect(url_for('index'))

@app.route('/students')
def students():
    """Student management page"""
    students = get_all_students()
    return render_template('students.html', students=students)

@app.route('/rooms')
def rooms():
    """Room management page"""
    rooms = get_all_rooms()
    return render_template('rooms.html', rooms=rooms)

@app.route('/generate_seating_page')
def generate_seating_page():
    """Dedicated seating generation page"""
    students = get_all_students()
    rooms = get_all_rooms()
    
    # Get unique classes for filtering
    classes = list(set([s['class'] for s in students]))
    classes.sort()
    
    return render_template('generate.html', classes=classes, student_count=len(students), room_count=len(rooms))

@app.route('/reports')
def reports():
    """Reports and analytics page"""
    plans = get_all_seating_plans()
    return render_template('reports.html', plans=plans)

@app.route('/add_student_page')
def add_student_page():
    """Add student form page"""
    return render_template('add_student.html')

@app.route('/edit_student_page/<roll_number>')
def edit_student_page(roll_number):
    """Edit student form page"""
    student = get_student_by_roll(roll_number)
    if not student:
        flash('Student not found', 'error')
        return redirect(url_for('students'))
    return render_template('edit_student.html', student=student)

@app.route('/add_room_page')
def add_room_page():
    """Add room form page"""
    return render_template('add_room.html')

@app.route('/edit_room_page/<room_id>')
def edit_room_page(room_id):
    """Edit room form page"""
    room = get_room_by_id(room_id)
    if not room:
        flash('Room not found', 'error')
        return redirect(url_for('rooms'))
    return render_template('edit_room.html', room=room)

@app.route('/update_student', methods=['POST'])
def update_student():
    """Update an existing student"""
    try:
        original_roll = request.form.get('original_roll_number', '').strip()
        new_roll = request.form.get('roll_number', '').strip()
        name = request.form.get('name', '').strip()
        class_name = request.form.get('class_name', '').strip()
        section = request.form.get('section', '').strip()
        gender = request.form.get('gender', '').strip().lower()
        
        if not all([original_roll, new_roll, name, class_name, section, gender]):
            flash('All fields are required', 'error')
            return redirect(url_for('edit_student_page', roll_number=original_roll))
        
        success = update_student(original_roll, new_roll, name, class_name, section, gender)
        if success:
            flash('Student updated successfully', 'success')
            return redirect(url_for('students'))
        else:
            flash('Roll number already exists', 'error')
            return redirect(url_for('edit_student_page', roll_number=original_roll))
        
    except Exception as e:
        logging.error(f"Error updating student: {str(e)}")
        flash('Error updating student', 'error')
        return redirect(url_for('students'))

@app.route('/delete_student', methods=['POST'])
def delete_student():
    """Delete a student"""
    try:
        roll_number = request.form.get('roll_number', '').strip()
        
        if not roll_number:
            flash('Roll number is required', 'error')
            return redirect(url_for('students'))
        
        delete_student(roll_number)
        flash('Student deleted successfully', 'success')
        return redirect(url_for('students'))
        
    except Exception as e:
        logging.error(f"Error deleting student: {str(e)}")
        flash('Error deleting student', 'error')
        return redirect(url_for('students'))

@app.route('/update_room', methods=['POST'])
def update_room():
    """Update an existing room"""
    try:
        original_room_id = request.form.get('original_room_id', '').strip()
        new_room_id = request.form.get('room_id', '').strip()
        rows = int(request.form.get('rows', 0))
        columns = int(request.form.get('columns', 0))
        extra_desks = int(request.form.get('extra_desks', 0))
        
        if not new_room_id or rows <= 0 or columns <= 0:
            flash('Invalid room configuration', 'error')
            return redirect(url_for('edit_room_page', room_id=original_room_id))
        
        success = update_room(original_room_id, new_room_id, rows, columns, extra_desks)
        if success:
            flash('Room updated successfully', 'success')
            return redirect(url_for('rooms'))
        else:
            flash('Room ID already exists', 'error')
            return redirect(url_for('edit_room_page', room_id=original_room_id))
        
    except ValueError:
        flash('Invalid number format', 'error')
        return redirect(url_for('edit_room_page', room_id=original_room_id))
    except Exception as e:
        logging.error(f"Error updating room: {str(e)}")
        flash('Error updating room', 'error')
        return redirect(url_for('rooms'))

@app.route('/delete_room', methods=['POST'])
def delete_room():
    """Delete a room"""
    try:
        room_id = request.form.get('room_id', '').strip()
        
        if not room_id:
            flash('Room ID is required', 'error')
            return redirect(url_for('rooms'))
        
        delete_room(room_id)
        flash('Room deleted successfully', 'success')
        return redirect(url_for('rooms'))
        
    except Exception as e:
        logging.error(f"Error deleting room: {str(e)}")
        flash('Error deleting room', 'error')
        return redirect(url_for('rooms'))

@app.route('/delete_plan/<int:plan_id>', methods=['POST'])
def delete_plan(plan_id):
    """Delete a seating plan"""
    try:
        delete_seating_plan(plan_id)
        flash('Seating plan deleted successfully', 'success')
    except Exception as e:
        logging.error(f"Error deleting seating plan: {str(e)}")
        flash('Error deleting seating plan', 'error')
    return redirect(url_for('reports'))

@app.route('/export_students', methods=['POST'])
def export_students():
    """Export students data to Excel"""
    try:
        students = get_all_students()
        if not students:
            flash('No student data to export', 'warning')
            return redirect(url_for('students'))
        
        # Create filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"students_export_{timestamp}.xlsx"
        filepath = os.path.join('exports', filename)
        
        # Export to Excel
        df = pd.DataFrame([dict(s) for s in students])
        df.to_excel(filepath, index=False, engine='openpyxl')
        
        flash(f'Students exported successfully: {filename}', 'success')
        return send_file(filepath, as_attachment=True, download_name=filename)
        
    except Exception as e:
        logging.error(f"Error exporting students: {str(e)}")
        flash('Error exporting students', 'error')
        return redirect(url_for('students'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)