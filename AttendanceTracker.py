# Import necessary libraries - pandas to handle Excel file and matplotlib for data visualization
import pandas as pd
import matplotlib.pyplot as plt

# Welcome message
print("Welcome to the Attendance Tracker!")

def create_excel_file(filename):
    # Create a new excel file and add headers if not exists
    try:
        pd.read_excel(filename)
        print(f"The excel file {filename} already exists!")
    except FileNotFoundError:
        # Create an empty Excel file with specified column headers
        df = pd.DataFrame(columns=["Employee Name", "Date", "Status"])
        df.to_excel(filename, index=False)
        print(f"The excel file {filename} was created successfully!")

def record_attendance(filename):
    name = input("Enter employee name: ")
    date = input("Enter date (YYYY-MM-DD): ")
    status = input("Enter status (Present, Absent, Late): ")

    # Append the new attendance record to the Excel file
    try:
        df = pd.read_excel(filename)
        new_record = {"Employee Name": name, "Date": date, "Status": status}
        df = df.append(new_record, ignore_index=True)
        df.to_excel(filename, index=False)
        print("Attendance recorded successfully.")
    except Exception as e:
        print(f"An error occurred while recording attendance: {e}")

def view_and_visualize_attendance_record(filename):
    try:
        df = pd.read_excel(filename)
        if df.empty:
            print("No attendance records found.")
            return
        else:
            print("Displaying Attendance Records:")
            print(df)

        attendance_summary = df['Status'].value_counts()

        plt.figure(figsize=(10, 6))
        attendance_summary.plot(kind='bar', color=['green', 'red', 'blue'])
        plt.title('Attendance Summary')
        plt.xlabel('Attendance Status')
        plt.ylabel('Number of Records')
        plt.xticks(rotation=0)
        plt.grid(True)
        plt.show()
    except FileNotFoundError:
        print("No attendance records found. Please record some attendance first.")
    except Exception as e:
        print(f"An error occurred while visualizing attendance: {e}")

def main():
    filename = 'attendance_records.xlsx'
    
    # Create Excel file if it doesn't exist
    create_excel_file(filename)
    
    while True:
        print("\nPlease choose an option:")
        print("1. Record attendance")
        print("2. View and Visualize attendance record")
        print("3. Quit")
        
        try:
            choice = int(input("Enter your choice: "))
        except ValueError:
            print("Invalid input. Please enter a number (1, 2 or 3).")
            continue

        if choice == 1:
            record_attendance(filename)
        elif choice == 2:
            view_and_visualize_attendance_record(filename)
        elif choice == 3:
            print("Thank you for using the Attendance Tracker. Goodbye!")
            break
        else:
            print("Invalid option!")

if __name__ == "__main__":
    main()
