# This application uses in-memory storage via Flask's app.config
# No database models are needed for the current requirements
# All data is stored in session/config and exported to Excel files

# Future enhancement: Add database models for persistent storage
# from app import db

# class Student(db.Model):
#     id = db.Column(db.Integer, primary_key=True)
#     roll_number = db.Column(db.String(50), unique=True, nullable=False)
#     name = db.Column(db.String(100), nullable=False)
#     class_name = db.Column(db.String(50), nullable=False)
#     section = db.Column(db.String(10), nullable=False)
#     gender = db.Column(db.String(10), nullable=False)

# class Room(db.Model):
#     id = db.Column(db.Integer, primary_key=True)
#     room_id = db.Column(db.String(50), unique=True, nullable=False)
#     rows = db.Column(db.Integer, nullable=False)
#     columns = db.Column(db.Integer, nullable=False)
