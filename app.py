import os
import logging
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
from seating_algorithm import SeatingAlgorithm
from excel_handler import ExcelHandler
import json

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

# Initialize handlers
excel_handler = ExcelHandler()
seating_algorithm = SeatingAlgorithm()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

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
            students = excel_handler.read_student_data(filepath)
            if students is not None:
                # Store in session for use in seating generation
                app.config['STUDENT_DATA'] = students.to_dict('records')
                flash(f'Successfully uploaded {len(students)} students', 'success')
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
        gender = request.form.get('gender', '').strip()
        
        if not all([roll_number, name, class_name, section, gender]):
            flash('All fields are required', 'error')
            return redirect(url_for('index'))
        
        # Initialize student data if not exists
        if 'STUDENT_DATA' not in app.config:
            app.config['STUDENT_DATA'] = []
        
        # Check for duplicate roll number
        existing_rolls = [s['roll_number'] for s in app.config['STUDENT_DATA']]
        if roll_number in existing_rolls:
            flash('Roll number already exists', 'error')
            return redirect(url_for('index'))
        
        # Add new student
        new_student = {
            'roll_number': roll_number,
            'name': name,
            'class': class_name,
            'section': section,
            'gender': gender
        }
        app.config['STUDENT_DATA'].append(new_student)
        flash('Student added successfully', 'success')
        
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
        
        # Initialize rooms data if not exists
        if 'ROOMS_DATA' not in app.config:
            app.config['ROOMS_DATA'] = []
        
        # Check for duplicate room ID
        existing_rooms = [r['room_id'] for r in app.config['ROOMS_DATA']]
        if room_id in existing_rooms:
            flash('Room ID already exists', 'error')
            return redirect(url_for('index'))
        
        # Calculate total capacity
        base_capacity = rows * columns * 2
        total_capacity = base_capacity + (extra_desks * 2)
        
        # Add new room
        new_room = {
            'room_id': room_id,
            'rows': rows,
            'columns': columns,
            'extra_desks': extra_desks,
            'base_capacity': base_capacity,
            'capacity': total_capacity
        }
        app.config['ROOMS_DATA'].append(new_room)
        flash(f'Room added successfully with {rows}×{columns} grid + {extra_desks} extra desks (Total: {total_capacity} students)', 'success')
        
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
        
        # Validate data
        if 'STUDENT_DATA' not in app.config or not app.config['STUDENT_DATA']:
            flash('No student data available. Please upload or add students first.', 'error')
            return redirect(url_for('index'))
        
        if 'ROOMS_DATA' not in app.config or not app.config['ROOMS_DATA']:
            flash('No rooms configured. Please add rooms first.', 'error')
            return redirect(url_for('index'))
        
        students = app.config['STUDENT_DATA']
        rooms = app.config['ROOMS_DATA']
        
        # Filter students if specific classes selected
        if selected_classes:
            students = [s for s in students if s['class'] in selected_classes]
        
        if not students:
            flash('No students match the selected criteria.', 'error')
            return redirect(url_for('index'))
        
        # Generate seating arrangement
        seating_plan, unassigned = seating_algorithm.generate_seating_plan(
            students, rooms, boy_girl_pairing, gender_separation, random_assignment
        )
        
        if not seating_plan:
            flash('Could not generate seating plan. Please check room capacity and constraints.', 'error')
            return redirect(url_for('index'))
        
        # Store results for preview and export
        app.config['SEATING_PLAN'] = seating_plan
        app.config['UNASSIGNED_STUDENTS'] = unassigned
        app.config['GENERATION_CONFIG'] = {
            'selected_classes': selected_classes,
            'boy_girl_pairing': boy_girl_pairing,
            'gender_separation': gender_separation,
            'random_assignment': random_assignment
        }
        
        flash(f'Seating plan generated successfully! {len(unassigned)} students remain unassigned.', 'success')
        return redirect(url_for('preview'))
        
    except Exception as e:
        logging.error(f"Error generating seating: {str(e)}")
        flash(f'Error generating seating plan: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/preview')
def preview():
    if 'SEATING_PLAN' not in app.config:
        flash('No seating plan available. Please generate one first.', 'error')
        return redirect(url_for('index'))
    
    seating_plan = app.config['SEATING_PLAN']
    unassigned = app.config.get('UNASSIGNED_STUDENTS', [])
    rooms = app.config.get('ROOMS_DATA', [])
    
    return render_template('preview.html', 
                         seating_plan=seating_plan, 
                         unassigned=unassigned, 
                         rooms=rooms)

@app.route('/export_room/<room_id>')
def export_room(room_id):
    try:
        if 'SEATING_PLAN' not in app.config:
            flash('No seating plan available', 'error')
            return redirect(url_for('index'))
        
        seating_plan = app.config['SEATING_PLAN']
        rooms = app.config.get('ROOMS_DATA', [])
        
        # Find the room
        room = next((r for r in rooms if r['room_id'] == room_id), None)
        if not room:
            flash('Room not found', 'error')
            return redirect(url_for('preview'))
        
        # Get seating data for this room
        room_seating = seating_plan.get(room_id, {})
        
        # Generate Excel file
        filename = excel_handler.export_room_seating(room_id, room, room_seating)
        if filename:
            return send_file(filename, as_attachment=True)
        else:
            flash('Error generating Excel file', 'error')
            return redirect(url_for('preview'))
            
    except Exception as e:
        logging.error(f"Error exporting room: {str(e)}")
        flash(f'Error exporting room: {str(e)}', 'error')
        return redirect(url_for('preview'))

@app.route('/get_student_data')
def get_student_data():
    students = app.config.get('STUDENT_DATA', [])
    return jsonify(students)

@app.route('/get_rooms_data')
def get_rooms_data():
    rooms = app.config.get('ROOMS_DATA', [])
    return jsonify(rooms)

@app.route('/load_sample_data', methods=['POST'])
def load_sample_data():
    try:
        # Load sample student data
        import pandas as pd
        
        # Try to load the JNV test data file if it exists
        if os.path.exists('jnv_students_test_data.xlsx'):
            df = pd.read_excel('jnv_students_test_data.xlsx')
            app.config['STUDENT_DATA'] = df.to_dict('records')
            
            # Add sample rooms
            sample_rooms = [
                {'room_id': 'EXAM-HALL-A', 'rows': 6, 'columns': 8, 'extra_desks': 2, 'base_capacity': 96, 'capacity': 100},
                {'room_id': 'EXAM-HALL-B', 'rows': 5, 'columns': 7, 'extra_desks': 3, 'base_capacity': 70, 'capacity': 76},
                {'room_id': 'LIBRARY-HALL', 'rows': 8, 'columns': 6, 'extra_desks': 0, 'base_capacity': 96, 'capacity': 96},
                {'room_id': 'COMPUTER-LAB', 'rows': 4, 'columns': 8, 'extra_desks': 2, 'base_capacity': 64, 'capacity': 68},
            ]
            app.config['ROOMS_DATA'] = sample_rooms
            
            flash(f'Sample data loaded: {len(df)} students and 4 exam halls configured', 'success')
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
        app.config.pop('STUDENT_DATA', None)
        flash('Student data cleared', 'info')
    elif data_type == 'rooms':
        app.config.pop('ROOMS_DATA', None)
        flash('Room data cleared', 'info')
    elif data_type == 'seating':
        app.config.pop('SEATING_PLAN', None)
        app.config.pop('UNASSIGNED_STUDENTS', None)
        app.config.pop('GENERATION_CONFIG', None)
        flash('Seating plan cleared', 'info')
    elif data_type == 'all':
        app.config.pop('STUDENT_DATA', None)
        app.config.pop('ROOMS_DATA', None)
        app.config.pop('SEATING_PLAN', None)
        app.config.pop('UNASSIGNED_STUDENTS', None)
        app.config.pop('GENERATION_CONFIG', None)
        flash('All data cleared', 'info')
    
    return redirect(url_for('index'))

@app.route('/students')
def students():
    """Student management page"""
    return render_template('students.html')

@app.route('/rooms')
def rooms():
    """Room management page"""
    return render_template('rooms.html')

@app.route('/generate_seating_page')
def generate_seating_page():
    """Dedicated seating generation page"""
    students_data = app.config.get('STUDENT_DATA', [])
    rooms_data = app.config.get('ROOMS_DATA', [])
    
    # Get unique classes for filtering
    classes = list(set([s['class'] for s in students_data if 'class' in s]))
    classes.sort()
    
    return render_template('generate.html', classes=classes, student_count=len(students_data), room_count=len(rooms_data))

@app.route('/reports')
def reports():
    """Reports and analytics page"""
    return render_template('reports.html')

@app.route('/add_student_page')
def add_student_page():
    """Add student form page"""
    return render_template('add_student.html')

@app.route('/edit_student_page')
def edit_student_page():
    """Edit student form page"""
    return render_template('edit_student.html')

@app.route('/add_room_page')
def add_room_page():
    """Add room form page"""
    return render_template('add_room.html')

@app.route('/edit_room_page')
def edit_room_page():
    """Edit room form page"""
    return render_template('edit_room.html')

@app.route('/update_student', methods=['POST'])
def update_student():
    """Update an existing student"""
    try:
        original_roll = request.form.get('original_roll_number', '').strip()
        new_roll = request.form.get('roll_number', '').strip()
        name = request.form.get('name', '').strip()
        class_name = request.form.get('class_name', '').strip()
        section = request.form.get('section', '').strip()
        gender = request.form.get('gender', '').strip()
        
        if not all([original_roll, new_roll, name, class_name, section, gender]):
            flash('All fields are required', 'error')
            return redirect(url_for('edit_student_page') + f'?roll_number={original_roll}')
        
        students_data = app.config.get('STUDENT_DATA', [])
        
        # Find the student to update
        student_found = False
        for i, student in enumerate(students_data):
            student_roll = student.get('roll_number') or student.get('Roll Number')
            if student_roll == original_roll:
                # Check if new roll number conflicts with another student
                if new_roll != original_roll:
                    existing_rolls = [s.get('roll_number') or s.get('Roll Number') for s in students_data if s != student]
                    if new_roll in existing_rolls:
                        flash('Roll number already exists', 'error')
                        return redirect(url_for('edit_student_page') + f'?roll_number={original_roll}')
                
                # Update student data
                students_data[i] = {
                    'roll_number': new_roll,
                    'name': name,
                    'class': f"{class_name}{section}",
                    'section': section,
                    'gender': gender.lower()
                }
                student_found = True
                break
        
        if not student_found:
            flash('Student not found', 'error')
            return redirect(url_for('students'))
        
        app.config['STUDENT_DATA'] = students_data
        flash('Student updated successfully', 'success')
        return redirect(url_for('students'))
        
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
        
        students_data = app.config.get('STUDENT_DATA', [])
        
        # Find and remove the student
        original_count = len(students_data)
        students_data = [s for s in students_data if 
                        (s.get('roll_number') or s.get('Roll Number')) != roll_number]
        
        if len(students_data) == original_count:
            flash('Student not found', 'error')
        else:
            app.config['STUDENT_DATA'] = students_data
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
            return redirect(url_for('edit_room_page') + f'?room_id={original_room_id}')
        
        rooms_data = app.config.get('ROOMS_DATA', [])
        
        # Find the room to update
        room_found = False
        for i, room in enumerate(rooms_data):
            if room['room_id'] == original_room_id:
                # Check if new room ID conflicts
                if new_room_id != original_room_id:
                    existing_rooms = [r['room_id'] for r in rooms_data if r != room]
                    if new_room_id in existing_rooms:
                        flash('Room ID already exists', 'error')
                        return redirect(url_for('edit_room_page') + f'?room_id={original_room_id}')
                
                # Calculate capacities
                base_capacity = rows * columns * 2
                total_capacity = base_capacity + (extra_desks * 2)
                
                # Update room data
                rooms_data[i] = {
                    'room_id': new_room_id,
                    'rows': rows,
                    'columns': columns,
                    'extra_desks': extra_desks,
                    'base_capacity': base_capacity,
                    'capacity': total_capacity
                }
                room_found = True
                break
        
        if not room_found:
            flash('Room not found', 'error')
            return redirect(url_for('rooms'))
        
        app.config['ROOMS_DATA'] = rooms_data
        flash('Room updated successfully', 'success')
        return redirect(url_for('rooms'))
        
    except ValueError:
        flash('Invalid number format', 'error')
        original_room_id = request.form.get('original_room_id', '')
        if original_room_id:
            return redirect(url_for('edit_room_page') + f'?room_id={original_room_id}')
        return redirect(url_for('rooms'))
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
        
        rooms_data = app.config.get('ROOMS_DATA', [])
        
        # Find and remove the room
        original_count = len(rooms_data)
        rooms_data = [r for r in rooms_data if r['room_id'] != room_id]
        
        if len(rooms_data) == original_count:
            flash('Room not found', 'error')
        else:
            app.config['ROOMS_DATA'] = rooms_data
            flash('Room deleted successfully', 'success')
        
        return redirect(url_for('rooms'))
        
    except Exception as e:
        logging.error(f"Error deleting room: {str(e)}")
        flash('Error deleting room', 'error')
        return redirect(url_for('rooms'))



@app.route('/get_seating_status')
def get_seating_status():
    """Get current seating plan status for reports"""
    return jsonify({
        'seating_plan': app.config.get('SEATING_PLAN', {}),
        'unassigned_students': app.config.get('UNASSIGNED_STUDENTS', []),
        'generation_config': app.config.get('GENERATION_CONFIG', {})
    })

@app.route('/export_students', methods=['POST'])
def export_students():
    """Export students data to Excel"""
    try:
        students_data = app.config.get('STUDENT_DATA', [])
        if not students_data:
            flash('No student data to export', 'warning')
            return redirect(url_for('students'))
        
        # Create filename with timestamp
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"students_export_{timestamp}.xlsx"
        filepath = os.path.join('exports', filename)
        
        # Export to Excel
        df = pd.DataFrame(students_data)
        df.to_excel(filepath, index=False, engine='openpyxl')
        
        flash(f'Students exported successfully: {filename}', 'success')
        return send_file(filepath, as_attachment=True, download_name=filename)
        
    except Exception as e:
        logging.error(f"Error exporting students: {str(e)}")
        flash('Error exporting students', 'error')
        return redirect(url_for('students'))

@app.route('/export_room_plan/<room_id>')
def export_room_plan(room_id):
    """Export individual room seating plan to Excel"""
    try:
        seating_plan = app.config.get('SEATING_PLAN', {})
        rooms_data = app.config.get('ROOMS_DATA', [])
        
        if room_id not in seating_plan:
            flash('Room seating plan not found', 'error')
            return redirect(url_for('preview'))
        
        room_info = next((r for r in rooms_data if r['room_id'] == room_id), {})
        room_seating = seating_plan[room_id]
        
        excel_handler = ExcelHandler()
        
        # Export standard format
        filepath = excel_handler.export_room_seating(room_id, room_info, room_seating)
        
        if filepath and os.path.exists(filepath):
            filename = os.path.basename(filepath)
            flash(f'Room {room_id} plan exported successfully', 'success')
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            flash('Error exporting room plan', 'error')
            return redirect(url_for('preview'))
            
    except Exception as e:
        logging.error(f"Error exporting room plan: {str(e)}")
        flash('Error exporting room plan', 'error')
        return redirect(url_for('preview'))

@app.route('/export_room_grid/<room_id>')
def export_room_grid(room_id):
    """Export individual room seating plan in grid format"""
    try:
        seating_plan = app.config.get('SEATING_PLAN', {})
        rooms_data = app.config.get('ROOMS_DATA', [])
        
        if room_id not in seating_plan:
            flash('Room seating plan not found', 'error')
            return redirect(url_for('preview'))
        
        room_info = next((r for r in rooms_data if r['room_id'] == room_id), {})
        room_seating = seating_plan[room_id]
        
        excel_handler = ExcelHandler()
        
        # Export grid format
        filepath = excel_handler.export_room_grid_layout(room_id, room_info, room_seating)
        
        if filepath and os.path.exists(filepath):
            filename = os.path.basename(filepath)
            flash(f'Room {room_id} grid layout exported successfully', 'success')
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            flash('Error exporting grid layout', 'error')
            return redirect(url_for('preview'))
            
    except Exception as e:
        logging.error(f"Error exporting grid layout: {str(e)}")
        flash('Error exporting grid layout', 'error')
        return redirect(url_for('preview'))

@app.route('/export_all_plans_zip')
def export_all_plans_zip():
    """Export all room seating plans as ZIP file with both standard and grid formats"""
    try:
        seating_plan = app.config.get('SEATING_PLAN', {})
        rooms_data = app.config.get('ROOMS_DATA', [])
        
        if not seating_plan:
            flash('No seating plan found', 'error')
            return redirect(url_for('preview'))
        
        excel_handler = ExcelHandler()
        
        # Export all rooms as ZIP
        zip_filepath = excel_handler.export_all_rooms_zip(seating_plan, rooms_data)
        
        if zip_filepath and os.path.exists(zip_filepath):
            filename = os.path.basename(zip_filepath)
            flash('Complete seating plan exported successfully', 'success')
            return send_file(zip_filepath, as_attachment=True, download_name=filename)
        else:
            flash('Error exporting complete plan', 'error')
            return redirect(url_for('preview'))
            
    except Exception as e:
        logging.error(f"Error exporting complete plan: {str(e)}")
        flash('Error exporting complete plan', 'error')
        return redirect(url_for('preview'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
