#!/usr/bin/env python3
"""
Create comprehensive test data for PM SHRI JNV GB NAGAR seating arrangement system.
"""
import pandas as pd
import random
from faker import Faker
import os

def create_jnv_test_data():
    """Create realistic test data for JNV GB Nagar."""
    fake = Faker('en_IN')  # Indian locale for better names
    
    # JNV typical class structure
    classes_data = [
        {'class': '6', 'sections': ['A', 'B'], 'students_per_section': 30},
        {'class': '7', 'sections': ['A', 'B'], 'students_per_section': 28},
        {'class': '8', 'sections': ['A', 'B'], 'students_per_section': 32},
        {'class': '9', 'sections': ['A', 'B'], 'students_per_section': 35},
        {'class': '10', 'sections': ['A', 'B'], 'students_per_section': 40},
        {'class': '11', 'sections': ['A', 'B'], 'students_per_section': 25},
        {'class': '12', 'sections': ['A', 'B'], 'students_per_section': 22},
    ]
    
    students_data = []
    roll_counter = 1
    
    for class_info in classes_data:
        class_num = class_info['class']
        sections = class_info['sections']
        students_per_section = class_info['students_per_section']
        
        for section in sections:
            for i in range(students_per_section):
                # Generate realistic Indian names
                gender = random.choice(['Male', 'Female'])
                if gender == 'Male':
                    first_name = fake.first_name_male()
                else:
                    first_name = fake.first_name_female()
                
                last_name = fake.last_name()
                full_name = f"{first_name} {last_name}"
                
                # Create roll number in JNV format
                roll_number = f"JNV{class_num}{section}{str(i+1).zfill(3)}"
                
                student = {
                    'roll_number': roll_number,
                    'name': full_name,
                    'class': f"{class_num}{section}",  # Combined class-section format
                    'gender': gender.lower(),  # lowercase for consistency
                    'raw_class': class_num,
                    'section': section
                }
                
                students_data.append(student)
                roll_counter += 1
    
    # Create DataFrame
    df = pd.DataFrame(students_data)
    
    # Shuffle to make it more realistic
    df = df.sample(frac=1).reset_index(drop=True)
    
    # Save to Excel
    output_file = 'jnv_students_test_data.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    # Statistics
    total_students = len(df)
    classes = sorted(df['raw_class'].unique())
    gender_stats = df['gender'].value_counts().to_dict()
    class_stats = df.groupby(['raw_class', 'section']).size().reset_index(name='Count')
    
    print(f"✅ JNV GB Nagar Test Data Created: '{output_file}'")
    print(f"📊 Total Students: {total_students}")
    print(f"📚 Classes: {', '.join(classes)}")
    print(f"👥 Gender Distribution: {gender_stats}")
    print(f"\n📋 Class-wise Distribution:")
    for _, row in class_stats.iterrows():
        print(f"   Class {row['raw_class']}{row['section']}: {row['Count']} students")
    
    return output_file, df

def create_sample_rooms():
    """Create sample room configurations for JNV."""
    rooms = [
        {'room_id': 'EXAM-HALL-A', 'rows': 6, 'columns': 8, 'extra_desks': 2},
        {'room_id': 'EXAM-HALL-B', 'rows': 5, 'columns': 7, 'extra_desks': 3},
        {'room_id': 'EXAM-HALL-C', 'rows': 4, 'columns': 6, 'extra_desks': 1},
        {'room_id': 'LIBRARY-HALL', 'rows': 8, 'columns': 6, 'extra_desks': 0},
        {'room_id': 'COMPUTER-LAB', 'rows': 4, 'columns': 8, 'extra_desks': 2},
        {'room_id': 'SCIENCE-LAB', 'rows': 3, 'columns': 5, 'extra_desks': 4},
    ]
    
    print(f"\n🏢 Sample Room Configurations:")
    for room in rooms:
        base_capacity = room['rows'] * room['columns'] * 2
        total_capacity = base_capacity + (room['extra_desks'] * 2)
        print(f"   {room['room_id']}: {room['rows']}×{room['columns']} + {room['extra_desks']} extra = {total_capacity} students")
    
    return rooms

if __name__ == "__main__":
    print("🎓 Creating Test Data for PM SHRI JNV GB NAGAR")
    print("=" * 50)
    
    try:
        # Install faker if not available
        output_file, df = create_jnv_test_data()
        sample_rooms = create_sample_rooms()
        
        print(f"\n✅ Test data ready!")
        print(f"📁 Upload '{output_file}' to test the system")
        print(f"🔧 Use the sample room configurations above")
        
    except ImportError:
        print("❌ Error: 'faker' library not installed")
        print("📦 Installing faker...")
        os.system("pip install faker")
        print("✅ Please run the script again")
    except Exception as e:
        print(f"❌ Error creating test data: {e}")