# importing required labraries
import csv
import pandas as pd

# creating the student class
class Student:
    def __init__(self, name, roll_number):
        self.name = name
        self.roll_number = roll_number
        self.grades = {}

# function to add the grades
    def add_grade(self, subject, grade):
        if 0 <= grade <= 100:
            self.grades[subject] = grade
        else:
            print("Invalid grade. Please enter a grade between 0 and 100.")

# function to calcluate grades
    def calculate_average(self):
        if self.grades:
            return sum(self.grades.values()) / len(self.grades)
        return 0

# function to display from above taken inputs
    def display_info(self):
        print(f"Name: {self.name}")
        print(f"Roll Number: {self.roll_number}")
        print("Grades:")
        for subject, grade in self.grades.items():
            print(f"{subject}: {grade}")
        print(f"Average Grade: {self.calculate_average():.2f}")

# Creating a class for student tracking

class StudentTracker:
    def __init__(self):
        self.students = {}

# Creating a functions for adding student, adding grades, view student, calculate average, save data
    def add_student(self, name, roll_number):
        if roll_number not in self.students:
            self.students[roll_number] = Student(name, roll_number)
        else:
            print("Student with this roll number already exists.")

    def add_grades(self, roll_number, subject, grade):
        if roll_number in self.students:
            self.students[roll_number].add_grade(subject, grade)
        else:
            print("Student not found.")

    def view_student_details(self, roll_number):
        if roll_number in self.students:
            self.students[roll_number].display_info()
        else:
            print("Student not found.")

    def calculate_average(self, roll_number):
        if roll_number in self.students:
            return self.students[roll_number].calculate_average()
        else:
            print("Student not found.")
            return 0

# function to save the data to csv file
    def save_to_csv(self, filename):
        with open(filename, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Roll Number", "Name", "Subjects", "Grades", "Average"])
            for student in self.students.values():
                subjects = ', '.join(student.grades.keys())
                grades = ', '.join(str(grade) for grade in student.grades.values())
                writer.writerow([student.roll_number, student.name, subjects, grades, student.calculate_average()])

# function to save the excel format
    def save_to_excel(self, filename):
        # Check if the file has an .xlsx extension, append it if not
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        
        # Prepare data for export
        data = {
            "Roll Number": [],
            "Name": [],
            "Subjects": [],
            "Grades": [],
            "Average": []
        }
        for student in self.students.values():
            data["Roll Number"].append(student.roll_number)
            data["Name"].append(student.name)
            data["Subjects"].append(', '.join(student.grades.keys()))
            data["Grades"].append(', '.join(str(grade) for grade in student.grades.values()))
            data["Average"].append(student.calculate_average())
        
        # Convert to DataFrame and save to Excel
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False)
        print(f"Data successfully saved to {filename}")


# function to display and take input from user
def main():
    tracker = StudentTracker()
    while True:
        print("\nMenu:")
        print("1. Add Student")
        print("2. Add Grades")
        print("3. View Student Details")
        print("4. Calculate Average Grade")
        print("5. Save Data to CSV")
        print("6. Save Data to excel")
        print("7. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            name = input("Enter student's name: ")
            roll_number = input("Enter roll number: ")
            tracker.add_student(name, roll_number)
            
        elif choice == '2':
            roll_number = input("Enter roll number: ")
            subject = input("Enter subject: ")
            grade = float(input("Enter grade: "))
            tracker.add_grades(roll_number, subject, grade)
        
        elif choice == '3':
            roll_number = input("Enter roll number: ")
            tracker.view_student_details(roll_number)
        
        elif choice == '4':
            roll_number = input("Enter roll number: ")
            average = tracker.calculate_average(roll_number)
            print(f"Average Grade: {average:.2f}")
        
        elif choice == '5':
            filename = input("Enter filename to save CSV (e.g., students.csv): ")
            tracker.save_to_csv(filename)
            print(f"Data saved to {filename}")
        
        elif choice == '6':
            filename = input("Enter filename to save excel (e.g., students.excel): ")
            tracker.save_to_excel(filename)
            print(f"Data saved to {filename}")
        
        elif choice == '7':
            break
        else:
            print("Invalid choice. Please try again.")

# calling the mian function
if __name__ == "__main__":
    main()
