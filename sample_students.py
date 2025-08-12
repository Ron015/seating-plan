#!/usr/bin/env python3
"""
Generate a sample Excel file with student data for testing the seating arrangement generator.
"""
import pandas as pd
import os

def create_sample_student_data():
    """Create sample student data for testing."""
    sample_data = [
        # Class 10A
        {'Roll Number': '10A001', 'Name': 'John Smith', 'Class': '10', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '10A002', 'Name': 'Emma Johnson', 'Class': '10', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '10A003', 'Name': 'Michael Brown', 'Class': '10', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '10A004', 'Name': 'Sarah Davis', 'Class': '10', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '10A005', 'Name': 'David Wilson', 'Class': '10', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '10A006', 'Name': 'Lisa Anderson', 'Class': '10', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '10A007', 'Name': 'Robert Taylor', 'Class': '10', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '10A008', 'Name': 'Jennifer White', 'Class': '10', 'Section': 'A', 'Gender': 'Female'},
        
        # Class 10B
        {'Roll Number': '10B001', 'Name': 'Christopher Lee', 'Class': '10', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '10B002', 'Name': 'Amanda Martinez', 'Class': '10', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '10B003', 'Name': 'James Garcia', 'Class': '10', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '10B004', 'Name': 'Ashley Rodriguez', 'Class': '10', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '10B005', 'Name': 'Daniel Hernandez', 'Class': '10', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '10B006', 'Name': 'Jessica Lopez', 'Class': '10', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '10B007', 'Name': 'Matthew Gonzalez', 'Class': '10', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '10B008', 'Name': 'Lauren Wilson', 'Class': '10', 'Section': 'B', 'Gender': 'Female'},
        
        # Class 11A
        {'Roll Number': '11A001', 'Name': 'Andrew Miller', 'Class': '11', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '11A002', 'Name': 'Stephanie Moore', 'Class': '11', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '11A003', 'Name': 'Joshua Jackson', 'Class': '11', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '11A004', 'Name': 'Megan Thomas', 'Class': '11', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '11A005', 'Name': 'Kevin Thompson', 'Class': '11', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '11A006', 'Name': 'Rachel Clark', 'Class': '11', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '11A007', 'Name': 'Tyler Lewis', 'Class': '11', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '11A008', 'Name': 'Hannah Walker', 'Class': '11', 'Section': 'A', 'Gender': 'Female'},
        
        # Class 11B
        {'Roll Number': '11B001', 'Name': 'Brandon Hall', 'Class': '11', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '11B002', 'Name': 'Samantha Allen', 'Class': '11', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '11B003', 'Name': 'Ryan Young', 'Class': '11', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '11B004', 'Name': 'Nicole King', 'Class': '11', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '11B005', 'Name': 'Justin Wright', 'Class': '11', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '11B006', 'Name': 'Brittany Scott', 'Class': '11', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '11B007', 'Name': 'Nathan Green', 'Class': '11', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '11B008', 'Name': 'Courtney Adams', 'Class': '11', 'Section': 'B', 'Gender': 'Female'},
        
        # Class 12A
        {'Roll Number': '12A001', 'Name': 'Eric Baker', 'Class': '12', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '12A002', 'Name': 'Crystal Nelson', 'Class': '12', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '12A003', 'Name': 'Aaron Carter', 'Class': '12', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '12A004', 'Name': 'Danielle Mitchell', 'Class': '12', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '12A005', 'Name': 'Jeremy Perez', 'Class': '12', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '12A006', 'Name': 'Vanessa Roberts', 'Class': '12', 'Section': 'A', 'Gender': 'Female'},
        {'Roll Number': '12A007', 'Name': 'Cody Turner', 'Class': '12', 'Section': 'A', 'Gender': 'Male'},
        {'Roll Number': '12A008', 'Name': 'Kimberly Phillips', 'Class': '12', 'Section': 'A', 'Gender': 'Female'},
        
        # Class 12B
        {'Roll Number': '12B001', 'Name': 'Trevor Campbell', 'Class': '12', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '12B002', 'Name': 'Morgan Parker', 'Class': '12', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '12B003', 'Name': 'Austin Evans', 'Class': '12', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '12B004', 'Name': 'Chelsea Edwards', 'Class': '12', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '12B005', 'Name': 'Caleb Collins', 'Class': '12', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '12B006', 'Name': 'Alexis Stewart', 'Class': '12', 'Section': 'B', 'Gender': 'Female'},
        {'Roll Number': '12B007', 'Name': 'Lucas Sanchez', 'Class': '12', 'Section': 'B', 'Gender': 'Male'},
        {'Roll Number': '12B008', 'Name': 'Jasmine Morris', 'Class': '12', 'Section': 'B', 'Gender': 'Female'},
    ]
    
    # Create DataFrame
    df = pd.DataFrame(sample_data)
    
    # Save to Excel file
    output_file = 'sample_students.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    print(f"Sample student data created in '{output_file}'")
    print(f"Total students: {len(df)}")
    print(f"Classes: {sorted(df['Class'].unique())}")
    print(f"Sections per class: {df.groupby('Class')['Section'].nunique().to_dict()}")
    print(f"Gender distribution: {df['Gender'].value_counts().to_dict()}")
    
    return output_file

if __name__ == "__main__":
    create_sample_student_data()