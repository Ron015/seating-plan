import sqlite3
from flask import g, current_app

def get_db():
    """Get a database connection"""
    if 'db' not in g:
        g.db = sqlite3.connect(current_app.config['DATABASE'])
        g.db.row_factory = sqlite3.Row
    return g.db

def close_db(e=None):
    """Close the database connection"""
    db = g.pop('db', None)
    if db is not None:
        db.close()

def init_db():
    """Initialize the database with required tables"""
    db = get_db()
    
    # Create students table
    db.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_number TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            class TEXT NOT NULL,
            section TEXT NOT NULL,
            gender TEXT NOT NULL
        )
    """)
    
    # Create rooms table
    db.execute("""
        CREATE TABLE IF NOT EXISTS rooms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            room_id TEXT UNIQUE NOT NULL,
            rows INTEGER NOT NULL,
            columns INTEGER NOT NULL,
            extra_desks INTEGER DEFAULT 0,
            base_capacity INTEGER NOT NULL,
            capacity INTEGER NOT NULL
        )
    """)
    
    # Create seating_plans table
    db.execute("""
        CREATE TABLE IF NOT EXISTS seating_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_name TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            config_json TEXT NOT NULL
        )
    """)
    
    # Create seating_assignments table
    db.execute("""
        CREATE TABLE IF NOT EXISTS seating_assignments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_id INTEGER NOT NULL,
            room_id TEXT NOT NULL,
            desk_number TEXT NOT NULL,
            student_id INTEGER NOT NULL,
            FOREIGN KEY (plan_id) REFERENCES seating_plans(id),
            FOREIGN KEY (student_id) REFERENCES students(id)
        )
    """)
    
    # Create unassigned_students table
    db.execute("""
        CREATE TABLE IF NOT EXISTS unassigned_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            FOREIGN KEY (plan_id) REFERENCES seating_plans(id),
            FOREIGN KEY (student_id) REFERENCES students(id)
        )
    """)
    
    db.commit()

def query_db(query, args=(), one=False):
    """Execute a query and return results"""
    db = get_db()
    cur = db.execute(query, args)
    rv = cur.fetchall()
    cur.close()
    return (rv[0] if rv else None) if one else rv

def execute_db(query, args=()):
    """Execute a query that doesn't return results"""
    db = get_db()
    db.execute(query, args)
    db.commit()

# Student operations
def get_all_students():
    return query_db("SELECT * FROM students ORDER BY class, roll_number")

def get_student_by_roll(roll_number):
    return query_db("SELECT * FROM students WHERE roll_number = ?", [roll_number], one=True)

def add_student(roll_number, name, class_name, section, gender):
    try:
        execute_db(
            "INSERT INTO students (roll_number, name, class, section, gender) VALUES (?, ?, ?, ?, ?)",
            (roll_number, name, class_name, section, gender)
        )
        return True
    except sqlite3.IntegrityError:
        return False

def update_student(original_roll, new_roll, name, class_name, section, gender):
    try:
        execute_db(
            """UPDATE students SET roll_number = ?, name = ?, class = ?, section = ?, gender = ? 
            WHERE roll_number = ?""",
            (new_roll, name, class_name, section, gender, original_roll)
        )
        return True
    except sqlite3.IntegrityError:
        return False

def delete_student(roll_number):
    execute_db("DELETE FROM students WHERE roll_number = ?", [roll_number])

# Room operations
def get_all_rooms():
    return query_db("SELECT * FROM rooms ORDER BY room_id")

def get_room_by_id(room_id):
    return query_db("SELECT * FROM rooms WHERE room_id = ?", [room_id], one=True)

def add_room(room_id, rows, columns, extra_desks):
    base_capacity = rows * columns * 2
    total_capacity = base_capacity + (extra_desks * 2)
    try:
        execute_db(
            """INSERT INTO rooms (room_id, rows, columns, extra_desks, base_capacity, capacity) 
            VALUES (?, ?, ?, ?, ?, ?)""",
            (room_id, rows, columns, extra_desks, base_capacity, total_capacity)
        )
        return True
    except sqlite3.IntegrityError:
        return False

def update_room(original_room_id, new_room_id, rows, columns, extra_desks):
    base_capacity = rows * columns * 2
    total_capacity = base_capacity + (extra_desks * 2)
    try:
        execute_db(
            """UPDATE rooms SET room_id = ?, rows = ?, columns = ?, extra_desks = ?, 
            base_capacity = ?, capacity = ? WHERE room_id = ?""",
            (new_room_id, rows, columns, extra_desks, base_capacity, total_capacity, original_room_id)
        )
        return True
    except sqlite3.IntegrityError:
        return False

def delete_room(room_id):
    execute_db("DELETE FROM rooms WHERE room_id = ?", [room_id])

# Seating plan operations
def create_seating_plan(plan_name, config_json):
    cur = get_db().execute(
        "INSERT INTO seating_plans (plan_name, config_json) VALUES (?, ?)",
        (plan_name, config_json)
    )
    plan_id = cur.lastrowid
    get_db().commit()
    cur.close()
    return plan_id

def add_seating_assignment(plan_id, room_id, desk_number, student_id):
    execute_db(
        "INSERT INTO seating_assignments (plan_id, room_id, desk_number, student_id) VALUES (?, ?, ?, ?)",
        (plan_id, room_id, desk_number, student_id)
    )

def add_unassigned_student(plan_id, student_id):
    execute_db(
        "INSERT INTO unassigned_students (plan_id, student_id) VALUES (?, ?)",
        (plan_id, student_id)
    )

def get_seating_plan(plan_id):
    plan = query_db("SELECT * FROM seating_plans WHERE id = ?", [plan_id], one=True)
    if not plan:
        return None
    
    assignments = query_db(
        "SELECT * FROM seating_assignments WHERE plan_id = ?", [plan_id]
    )
    unassigned = query_db(
        "SELECT s.* FROM unassigned_students us JOIN students s ON us.student_id = s.id WHERE us.plan_id = ?",
        [plan_id]
    )
    
    return {
        'plan_info': plan,
        'assignments': assignments,
        'unassigned': unassigned
    }

def get_latest_seating_plan():
    return query_db(
        "SELECT * FROM seating_plans ORDER BY created_at DESC LIMIT 1", 
        one=True
    )

def get_room_assignments(plan_id, room_id):
    return query_db(
        """SELECT sa.desk_number, s.* 
        FROM seating_assignments sa 
        JOIN students s ON sa.student_id = s.id 
        WHERE sa.plan_id = ? AND sa.room_id = ?""",
        (plan_id, room_id)
    )

def get_all_seating_plans():
    return query_db("SELECT * FROM seating_plans ORDER BY created_at DESC")