#Matt Gurgiolo COOT Trip Sorter

import pandas as pd
import random

# Load the Excel file
xlsx_file = "2023 COOT Student data.xlsx"
xl = pd.ExcelFile(xlsx_file)

# Load the sheets into data frames
trip_sheet = xl.parse('Trip Sheet')
student_sheet = xl.parse('Student Sheet')

# Create a mapping from subcategories to trips
subcategory_to_trips = trip_sheet.groupby('Subcategory')['Trip'].apply(list).to_dict()

# Let's start by defining the Student and Trip classes

class Student:
    def __init__(self, first_name, last_name, id_number, preferences, gender, poc, dorm, team):
        self.first_name = first_name
        self.last_name = last_name
        self.id_number = id_number
        self.preferences = preferences  # Dictionary of {subcategory: score}
        self.gender = gender
        self.poc = poc
        self.dorm = dorm
        self.team = team
        self.assigned_trip = None
        
    def assign_trip(self, trip):
        self.assigned_trip = trip
        
    def preferences_str(self):
        str = self.preferences[0]
        for i in range(len(self.preferences)-1):
            str+= ", "+self.preferences[i+1]
        return str
        
    def toString(self):
        str = self.first_name+" "+self.last_name+"|| Preferences: "
        for i in range(len(self.preferences)):
            str+= self.preferences[i] + ", "
        return str


class Trip:
    def __init__(self, name, category, subcategory, capacity):
        self.name = name
        self.category = category
        self.subcategory = subcategory
        self.capacity = capacity
        self.assigned_students = []
        self.gender_distribution = {'M': 0, 'F': 0, 'Other': 0}
        self.poc_distribution = 0
        self.dorm_distribution = set()
        self.team_distribution = set()

    def add_student(self, student):
        self.assigned_students.append(student)
        self.capacity -= 1
        student.assign_trip(self)
        # Update gender distribution
        self.gender_distribution[student.gender] += 1
        # Update POC distribution
        if student.poc == 'Yes': self.poc_distribution += 1
        # Update dorm room distribution
        self.dorm_distribution.add(student.dorm)
        # Update team Distrubution
        if not student.team == 'N': self.team_distribution.add(student.team)

    def is_full(self):
        return self.capacity <= 0

    def get_assigned_students(self):
        return [student.id_number for student in self.assigned_students]
    
    def check_gender_distribution(self, student):
        # Calculate the current gender counts for M, F, and Other
        current_m = self.gender_distribution['M']
        current_f = self.gender_distribution['F']
        current_other = self.gender_distribution['Other']
        
        # Treat 'Other' as the same as the smaller gender
        if current_m <= current_f:
            current_m += current_other
        else:
            current_f += current_other
        
        # Update the current counts if the student is to be added
        if student.gender == 'M':
            future_m = current_m + 1
            future_f = current_f
        elif student.gender == 'F':
            future_m = current_m
            future_f = current_f + 1
        else:  # 'Other' is treated as the smaller gender
            future_m, future_f = (current_m + 1, current_f) if current_m <= current_f else (current_m, current_f + 1)
        
        # Calculate the future gender ratios if we add this student
        total_future = future_m + future_f
        future_ratio_m = future_m / total_future
        future_ratio_f = future_f / total_future
        if future_m == 0:
            future_m = .1
        if future_f == 0:
            future_f = .1
            
        if ((len(students)/2 + future_m)/future_f) > .8 or ((len(students)/2 + future_f)/future_m) > .8:
            return True
        # Check whether the future gender ratios are within 10% (0.1) of a 50-50 distribution
        return abs(future_ratio_m - 0.5) <= 0.1 and abs(future_ratio_f - 0.5) <= 0.1


    def check_poc_distribution(self):
        return self.poc_distribution != 1

    def check_dorm_distribution(self, dorm_room):
        return not self.dorm_distribution.__contains__(dorm_room)
    
    def check_team_distribution(self, team):
        if team == "N":
            return True
        return not self.team_distribution.__contains__(team)
    
    def print_students(self):
        string = self.name+ ":  "
        for s in self.assigned_students:
            string+=s.first_name
            string+=" "
            string+=s.last_name
            string+=", "
        print(string)
        
def add_student_check(trip, student):
        if trip.is_full():
            #print(f"Trip {trip.name} is full")
            return False
        if not trip.check_gender_distribution(student):
            #print(f"Adding student {student.first_name} {student.last_name} would disrupt gender distribution in trip {trip.name}")
            return False
        if not trip.check_dorm_distribution(student.dorm):
            #print(f"Adding student {student.first_name} {student.last_name} would disrupt dorm distribution in trip {trip.name}")
            return False
        if not trip.check_team_distribution(student.team):
            #print(f"Adding student {student.first_name} {student.last_name} would disrupt team distribution in trip {trip.name}")
            return False
        return True
        

#STARTING THE REPITIONS
students = []
trips = []
unassigned_students = []
assigned_students = []

last_first_percent = 0
for x in range(0, 10):
    students.clear()
    trips.clear()
    unassigned_students.clear()
    assigned_students.clear()

    #NOW CREATING THE ACTUAL OBJECTS 

    # Create Student objects from the 'Student Scores' DataFrame
    
    for _, row in student_sheet.iterrows():
        first_name = row['First Name']
        last_name = row['Last Name']
        id_number = row['Colby ID Number']
        gender = row['Gender']
        poc_status = row['POC']
        dorm_room = row['Dorm']
        team = row['Team']
        # Get the trip preferences and scores for this student
        preferences_dict = row[7:].dropna().to_dict()  # Exclude the non-score columns and any NaN scores
        # Trim everything before "-" and the space after in the subcategory names
        trimmed_preferences_dict = {k.split(' - ')[-1]: v for k, v in preferences_dict.items()}
        # Sort preferences by score in descending order and retain only the trimmed subcategory names
        preferences = sorted(trimmed_preferences_dict, key=trimmed_preferences_dict.get, reverse=True)
        # Create a Student object and add it to the list of students
        students.append(Student(first_name, last_name, id_number, preferences, gender, poc_status, dorm_room, team))
        
    random.shuffle(students)


    # Create Trip objects from the 'Trip Categories' DataFrame
    
    for _, row in trip_sheet.iterrows():
        name = row['Trip']
        category = row['Category']
        subcategory = row['Subcategory']
        capacity = row['Capacity']
        # Create a Trip object and add it to the list of trips
        trips.append(Trip(name, category, subcategory, capacity))

    #NOW SORTING STUDENTs


    choice_assigned_dict = {"First":0, "Second":0, "Third":0, "Fourth":0, "Fifth":0}

    for s in students:
        random.shuffle(trips)
        student_added = False
        for i in range(0, len(s.preferences)):
            subcat=s.preferences[i]
            for t in trips:
                if t.subcategory == subcat and add_student_check(t, s):
                    t.add_student(s) 
                    student_added = True
                    
                    if i==0: choice_assigned_dict["First"] += 1 
                    
                    if i==1: choice_assigned_dict["Second"] += 1
                    
                    if i==2: choice_assigned_dict["Third"] += 1
                    
                    if i==3: choice_assigned_dict["Fourth"] += 1
                    
                    if i==4: choice_assigned_dict["Fifth"] += 1
                    
                    #print(f"Student {s.first_name} {s.last_name} added to trip {t.name}")  # Print when a student is added
                    break  # Break the loop after a student is added
            if student_added:
                assigned_students.append(s) 
                break
        if not student_added: unassigned_students.append(s)
        

    ##THIS IS WRITING THE RESULTS TO A FILE
                        
    # Import the necessary library
    from pandas import ExcelWriter

    # Create a DataFrame for the "trips" sheet
    trips_data = {
        'Name': [trip.name for trip in trips],
        'Category': [trip.category for trip in trips],
        'Subcategory': [trip.subcategory for trip in trips],
        'Capacity': [trip.capacity for trip in trips],
        'Assigned Students': [", ".join(map(str, trip.get_assigned_students())) for trip in trips]
    }
    trips_df = pd.DataFrame(trips_data)

    # Create a DataFrame for the "assigned_students" sheet
    assigned_students_data = {
        'First Name': [student.first_name for student in assigned_students],
        'Last Name': [student.last_name for student in assigned_students],
        'ID Number': [student.id_number for student in assigned_students],
        'Preferences': [student.preferences_str() for student in assigned_students],
        'Gender': [student.gender for student in assigned_students],
        'POC': [student.poc for student in assigned_students],
        'Dorm': [student.dorm for student in assigned_students],
        'Team': [student.team for student in assigned_students],
        'Assigned Trip': [student.assigned_trip.name if student.assigned_trip else None for student in assigned_students]
    }
    assigned_students_df = pd.DataFrame(assigned_students_data)

    # Create a DataFrame for the "unassigned_students" sheet
    unassigned_students_data = {
        'First Name': [student.first_name for student in unassigned_students],
        'Last Name': [student.last_name for student in unassigned_students],
        'ID Number': [student.id_number for student in unassigned_students],
        'Preferences': [student.preferences_str() for student in unassigned_students],
        'Gender': [student.gender for student in unassigned_students],
        'POC': [student.poc for student in unassigned_students],
        'Dorm': [student.dorm for student in unassigned_students],
        'Team': [student.team for student in unassigned_students],
        'Assigned Trip': [student.assigned_trip.name if student.assigned_trip else None for student in unassigned_students]
    }
    unassigned_students_df = pd.DataFrame(unassigned_students_data)

    first_percent = choice_assigned_dict["First"]*100/len(assigned_students)
    if first_percent > last_first_percent:
        # Write the DataFrames to an Excel file
        with ExcelWriter('Sorter Results.xlsx') as writer:
            trips_df.to_excel(writer, sheet_name='trips', index=False)
            assigned_students_df.to_excel(writer, sheet_name='assigned_students', index=False)
            unassigned_students_df.to_excel(writer, sheet_name='unassigned_students', index=False)

        print("\nTrial "+ str(x) +" was successful. Here's the stats:")
        print("Percent of students assigned their First Choice: "+ str(choice_assigned_dict["First"]*100/len(assigned_students))+"%")
        print("Percent of students assigned their Second Choice: "+ str(choice_assigned_dict["Second"]*100/len(assigned_students))+"%")
        print("Percent of students assigned their Third Choice: "+ str(choice_assigned_dict["Third"]*100/len(assigned_students))+"%")
        print("Percent of students assigned their Fourth Choice: "+ str(choice_assigned_dict["Fourth"]*100/len(assigned_students))+"%")
        print("Percent of students assigned their Fifth Choice: "+ str(choice_assigned_dict["Fifth"]*100/len(assigned_students))+"%\n")
        
        last_first_percent = first_percent
                