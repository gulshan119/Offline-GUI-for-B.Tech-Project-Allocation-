#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def read_student_preferences(file_path):
    df = pd.read_excel(file_path)
    student_preferences = []
    for index, row in df.iterrows():
        student = row.iloc[0]
        preferences = row.iloc[1:].dropna().tolist()
        student_preferences.append({'student': student, 'preferences': preferences})
    return student_preferences

def read_faculty_preferences(file_path):
    df = pd.read_excel(file_path)
    faculty_order = df.iloc[:, 0].tolist()
    faculty_preferences = {}
    for index, row in df.iterrows():
        faculty = row.iloc[0]
        preferences = row.iloc[1:].dropna().tolist()
        faculty_preferences[faculty] = preferences
    return faculty_preferences, faculty_order

def allocate_students(student_preferences, faculty_preferences):
    allocated_students = {faculty: [] for faculty in faculty_preferences.keys()}
    all_students = student_preferences[:]
    preference_level = 0
    max_students_per_faculty = 8
    while all_students and preference_level < 5:
        students_left = []
        for student_data in all_students:
            student = student_data['student']
            if preference_level < len(student_data['preferences']):
                preferred_faculty = student_data['preferences'][preference_level]
                if preferred_faculty in faculty_preferences:
                    if len(faculty_preferences[preferred_faculty]) > 0 and len(allocated_students[preferred_faculty]) < max_students_per_faculty:
                        faculty_first_preference = faculty_preferences[preferred_faculty][0]
                        if student == faculty_first_preference:
                            allocated_students[preferred_faculty].append(student)
                            for faculty in faculty_preferences:
                                faculty_preferences[faculty] = [x for x in faculty_preferences[faculty] if x != student]
                            continue
            students_left.append(student_data)
        all_students = students_left
        preference_level += 1
    return allocated_students

def save_allocations_to_excel(allocations, faculty_order):
    allocations_data = {'Faculty': faculty_order, 'Students': []}
    for faculty in faculty_order:
        students = allocations.get(faculty, [])
        # Convert roll numbers to strings before joining
        student_strings = [str(student) for student in students]
        allocations_data['Students'].append(', '.join(student_strings))
    allocations_df = pd.DataFrame(allocations_data)
    allocations_file = 'allocations.xlsx'
    allocations_df.to_excel(allocations_file, index=False)
    return allocations_file

def save_unallocated_students_to_excel(unallocated_students):
    unallocated_df = pd.DataFrame({'Unallocated Students': unallocated_students})
    unallocated_file = 'unallocated_students.xlsx'
    unallocated_df.to_excel(unallocated_file, index=False)
    return unallocated_file

def upload_student_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        global student_preferences
        student_preferences = read_student_preferences(file_path)
        messagebox.showinfo("Success", "Student file uploaded successfully.")

def upload_faculty_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        global faculty_preferences, faculty_order
        faculty_preferences, faculty_order = read_faculty_preferences(file_path)
        messagebox.showinfo("Success", "Faculty file uploaded successfully.")

def allocate():
    if not student_preferences or not faculty_preferences:
        messagebox.showwarning("Warning", "Please upload both student and faculty preferences.")
        return
    
    allocated_students = allocate_students(student_preferences, faculty_preferences)
    
    all_students_set = {student['student'] for student in student_preferences}
    allocated_students_set = set(sum(allocated_students.values(), []))
    unallocated_students = list(all_students_set - allocated_students_set)
    
    allocations_file = save_allocations_to_excel(allocated_students, faculty_order)
    unallocated_file = save_unallocated_students_to_excel(unallocated_students)
    
    messagebox.showinfo("Success", f"Allocation completed.\nAllocations saved to: {allocations_file}\nUnallocated students saved to: {unallocated_file}")

if __name__ == '__main__':
    student_preferences = []
    faculty_preferences = {}
    faculty_order = []

    root = tk.Tk()
    root.title("BTP Allocation System")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    upload_student_button = tk.Button(frame, text="Upload Student Preferences", command=upload_student_file)
    upload_student_button.pack(pady=5)

    upload_faculty_button = tk.Button(frame, text="Upload Faculty Preferences", command=upload_faculty_file)
    upload_faculty_button.pack(pady=5)

    allocate_button = tk.Button(frame, text="Allocate", command=allocate)
    allocate_button.pack(pady=5)

    root.mainloop()
