import tkinter as tk
from tkinter import ttk, messagebox
import pypyodbc

class CollegeManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("UiTM Kedah InCoMS")

        # Create a StringVar for the student ID entry
        self.student_id_var = tk.StringVar()
        self.student_password_var = tk.StringVar()
        self.new_id_var = tk.StringVar()
        self.new_name_var = tk.StringVar()
        self.new_contact_var = tk.StringVar()
        self.new_faculty_var = tk.StringVar()
        self.new_gender_var = tk.StringVar()
        self.new_college_var = tk.StringVar()
        self.new_password_var = tk.StringVar()
        self.new_status_var = tk.StringVar()
        self.up_id_var = tk.StringVar()
        self.up_name_var = tk.StringVar()
        self.up_contact_var = tk.StringVar()
        self.up_faculty_var = tk.StringVar()
        self.up_gender_var = tk.StringVar()
        self.up_college_var = tk.StringVar()
        self.up_password_var = tk.StringVar()
        self.up_status_var = tk.StringVar()

        # Create widgets
        self.login_widgets()

    def login_widgets(self):
        # Title Label
        title_label = ttk.Label(self.root, text="Welcome to UiTM Kedah Integrated College Management System", font=("Arial", 15))
        title_label.grid(row=0, column=1, columnspan=4, padx=20, pady=20)

        # ID Label and Entry
        id_label = ttk.Label(self.root, text="Student/ Staff ID", font=("Arial", 12))
        id_label.grid(row=1, column=1, columnspan=1, padx=5, pady=5, sticky='e')
        id_entry = ttk.Entry(self.root, textvariable=self.student_id_var, width=20, font=("Arial", 12))
        id_entry.grid(row=1, column=2, columnspan=3, padx=5, pady=5, sticky='w')

        # Password Label and Entry
        pswrd_label = ttk.Label(self.root, text="Password", font=("Arial", 12))
        pswrd_label.grid(row=2, column=1, columnspan=1, padx=5, pady=5, sticky='e')
        pwsrd_entry = ttk.Entry(self.root, textvariable=self.student_password_var, width=20, font=("Arial", 12), show="*")
        pwsrd_entry.grid(row=2, column=2, columnspan=3, padx=5, pady=5, sticky='w')

        # Login Button
        login_button = ttk.Button(self.root, text="Login", command=self.login_action)
        login_button.grid(row=3, column=1, columnspan=4, padx=5, pady=5, ipadx=5, ipady=5)

        # Note Label
        note_label = ttk.Label(self.root, text="If you encounter any difficulties, please contact Information Technology Center", font=("Arial", 10))
        note_label.grid(row=4, column=1, columnspan=4, padx=10, pady=15)

    def login_action(self):
        # Database connection details
        connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\SCSM11\Downloads\Telegram Desktop\Project_Finalize\Project_Finalize\registration database.accdb;'
        user_ID = self.student_id_var.get()
        password = self.student_password_var.get()

        try:
            # Connect to the Access database
            conn = pypyodbc.connect(connection_string)
            cursor = conn.cursor()

            # Query the database to check if the user ID and password match
            cursor.execute('SELECT [UserID], [Status] FROM [User] WHERE UserID = ? AND Password = ?', (user_ID, password))
            user_data = cursor.fetchone()

            if user_data:
                user_id, user_status = user_data

                if user_status == 'Student':
                    self.navigate_to_student_page(user_id)
                elif user_status == 'Staff':
                    self.navigate_to_staff_page(user_id)
                else:
                    messagebox.showerror("Login Failed", "Invalid user role") 
            else:
                messagebox.showerror("Login Failed", "Invalid user ID or password") 

        except pypyodbc.Error as e:
            messagebox.showerror("Database Error", f"Error connecting to the database: {e}")

        finally:
            conn.close()

    def navigate_to_student_page(self, user_id):
        # Create a new window for the student page
        student_page_window = tk.Toplevel(self.root)
        student_page_window.title("UiTM Kedah InCoMS | Student Page")

        # Retrieve and display data for the student with the given user_id
        student_data = self.retrieve_student_data(user_id)

        # Create and display widgets to show student data
        self.create_student_page_widgets(student_page_window, student_data)

    def retrieve_student_data(self, user_id):
        # Query the database to retrieve data for the student with the given user_id
        connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\SCSM11\Downloads\Telegram Desktop\Project_Finalize\Project_Finalize\registration database.accdb;'
        try:
            conn = pypyodbc.connect(connection_string)
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM [User] WHERE UserID = ?', (user_id,))
            student_data = cursor.fetchone()

            return student_data

        except pypyodbc.Error as e:
            messagebox.showerror("Database Error", f"Error retrieving student data: {e}")

        finally:
            conn.close()

    def navigate_to_staff_page(self, user_id):
        # Create a new window for the student page
        staff_page_window = tk.Toplevel(self.root)
        staff_page_window.title("UiTM Kedah InCoMS | Staff Page")

        # Retrieve and display data for the student with the given user_id
        staff_data = self.retrieve_staff_data(user_id)

        # Create and display widgets to show student data
        self.create_staff_page_widgets(staff_page_window, staff_data)

    def retrieve_staff_data(self, user_id):
        # Query the database to retrieve data for the student with the given user_id
        connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\SCSM11\Downloads\Telegram Desktop\Project_Finalize\Project_Finalize\registration database.accdb;'
        try:
            conn = pypyodbc.connect(connection_string)
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM [User] WHERE UserID = ?', (user_id,))
            student_data = cursor.fetchone()

            return student_data

        except pypyodbc.Error as e:
            messagebox.showerror("Database Error", f"Error retrieving staff data: {e}")

        finally:
            conn.close()

    def create_student_page_widgets(self, window_stdn, student_data):
        ttk.Label(self.root, text="UiTM Kedah Integrated College Management System | Student Database")

        window_label = ttk.Label(window_stdn, text="Welcome to UiTM Kedah InCoMS (Student Database Viewer)", font=("Arial", 14))
        window_label.grid(row=0, column=1, columnspan=4, padx=20, pady=20)

        stdn_id_label = ttk.Label(window_stdn, text="Student ID", font=("Arial", 12))
        stdn_id_label.grid(row=1, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        stdn_id_data = ttk.Label(window_stdn, text=f"{student_data[1]}", font=("Arial", 12))
        stdn_id_data.grid(row=1, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        stdn_name_label = ttk.Label(window_stdn, text="Name", font=("Arial", 12))
        stdn_name_label.grid(row=2, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        stdn_name_data = ttk.Label(window_stdn, text=f"{student_data[4]}", font=("Arial", 12))
        stdn_name_data.grid(row=2, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        stdn_contact_label = ttk.Label(window_stdn, text="Contact Number", font=("Arial", 12))
        stdn_contact_label.grid(row=3, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        stdn_contact_data = ttk.Label(window_stdn, text=f"{student_data[5]}", font=("Arial", 12))
        stdn_contact_data.grid(row=3, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        stdn_faculty_label = ttk.Label(window_stdn, text="Faculty", font=("Arial", 12))
        stdn_faculty_label.grid(row=4, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        stdn_faculty_data = ttk.Label(window_stdn, text=f"{student_data[7]}", font=("Arial", 12))
        stdn_faculty_data.grid(row=4, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        stdn_gender_label = ttk.Label(window_stdn, text="Gender", font=("Arial", 12))
        stdn_gender_label.grid(row=5, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        stdn_gender_data = ttk.Label(window_stdn, text=f"{student_data[6]}", font=("Arial", 12))
        stdn_gender_data.grid(row=5, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        stdn_college_label = ttk.Label(window_stdn, text="College Name", font=("Arial", 12))
        stdn_college_label.grid(row=6, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        stdn_college_data = ttk.Label(window_stdn, text=f"{student_data[8]}", font=("Arial", 12))
        stdn_college_data.grid(row=6, column=2, columnspan=3, padx=15, pady=5, sticky='w')

    def create_staff_page_widgets(self, window_staff, staff_data):
        ttk.Label(self.root, text="UiTM Kedah Integrated College Management System | Staff Database Manager")

        window_label = ttk.Label(window_staff, text="Welcome to UiTM Kedah InCoMS (Staff Database Manager)", font=("Arial", 14))
        window_label.grid(row=0, column=1, columnspan=4, padx=20, pady=20)

        staff_id_label = ttk.Label(window_staff, text="Staff ID", font=("Arial", 12))
        staff_id_label.grid(row=1, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        staff_id_data = ttk.Label(window_staff, text=f"{staff_data[1]}", font=("Arial", 12))
        staff_id_data.grid(row=1, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        staff_name_label = ttk.Label(window_staff, text="Name", font=("Arial", 12))
        staff_name_label.grid(row=2, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        staff_name_data = ttk.Label(window_staff, text=f"{staff_data[4]}", font=("Arial", 12))
        staff_name_data.grid(row=2, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        staff_contact_label = ttk.Label(window_staff, text="Contact Number", font=("Arial", 12))
        staff_contact_label.grid(row=3, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        staff_contact_data = ttk.Label(window_staff, text=f"{staff_data[5]}", font=("Arial", 12))
        staff_contact_data.grid(row=3, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        staff_contact_label = ttk.Label(window_staff, text="Faculty", font=("Arial", 12))
        staff_contact_label.grid(row=4, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        staff_contact_data = ttk.Label(window_staff, text=f"{staff_data[7]}", font=("Arial", 12))
        staff_contact_data.grid(row=4, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        staff_contact_label = ttk.Label(window_staff, text="Gender", font=("Arial", 12))
        staff_contact_label.grid(row=5, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        staff_contact_data = ttk.Label(window_staff, text=f"{staff_data[6]}", font=("Arial", 12))
        staff_contact_data.grid(row=5, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        staff_contact_label = ttk.Label(window_staff, text="College Name", font=("Arial", 12))
        staff_contact_label.grid(row=6, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        staff_contact_data = ttk.Label(window_staff, text=f"{staff_data[8]}", font=("Arial", 12))
        staff_contact_data.grid(row=6, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        button_frame = tk.Frame(window_staff)
        button_frame.grid(row=7, column=1, columnspan=4, padx=5, pady=5)

        staff_add_button = ttk.Button(button_frame, text="Create new data", command=self.new_data_page)
        staff_add_button.grid(row=1, column=1, columnspan=1, padx=15, pady=15, ipadx=5, ipady=5, sticky='e')

        staff_edit_button = ttk.Button(button_frame, text="Edit exsisting data", command=self.edit_data_page)
        staff_edit_button.grid(row=1, column=2, columnspan=1, padx=15, pady=15, ipadx=5, ipady=5, sticky='w')
    
    def new_data_page(self):
        # Create a new window for new data page
        new_page_window = tk.Toplevel(self.root)
        new_page_window.title("UiTM Kedah InCoMS | New Data Registration Page")

        window_label = ttk.Label(new_page_window, text="New Data Registration", font=("Arial", 14))
        window_label.grid(row=0, column=1, columnspan=4, padx=20, pady=20)

        new_id = ttk.Label(new_page_window, text="Student ID", font=("Arial", 12))
        new_id.grid(row=1, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_id_data = ttk.Entry(new_page_window, textvariable=self.new_id_var, width=30, font=("Arial", 12))
        new_id_data.grid(row=1, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_name = ttk.Label(new_page_window, text="Name", font=("Arial", 12))
        new_name.grid(row=2, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_name_data = ttk.Entry(new_page_window, textvariable=self.new_name_var, width=30, font=("Arial", 12))
        new_name_data.grid(row=2, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_contact = ttk.Label(new_page_window, text="Contact Number", font=("Arial", 12))
        new_contact.grid(row=3, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_contact_data = ttk.Entry(new_page_window, textvariable=self.new_contact_var, width=30, font=("Arial", 12))
        new_contact_data.grid(row=3, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_faculty = ttk.Label(new_page_window, text="Faculty", font=("Arial", 12))
        new_faculty.grid(row=4, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_faculty_data = ttk.Combobox(new_page_window, textvariable=self.new_faculty_var, width=28, font=("Arial", 12), values=('', 'Bussines', 'Engineering', 'Finance', 'Science', 'Pharmacy', 'Mathematics'))
        new_faculty_data.grid(row=4, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_gender = ttk.Label(new_page_window, text="Gender", font=("Arial", 12))
        new_gender.grid(row=5, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_gender_data = ttk.Combobox(new_page_window, textvariable=self.new_gender_var, width=28, font=("Arial", 12), values=('', 'Male', 'Female'))
        new_gender_data.grid(row=5, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_college = ttk.Label(new_page_window, text="College", font=("Arial", 12))
        new_college.grid(row=6, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_college_data = ttk.Combobox(new_page_window, textvariable=self.new_college_var, width=28, font=("Arial", 12), values=('', 'Kolej Mahsuri', 'Kolej Murni', 'Kolej Masria', 'Kolej Malinja'))
        new_college_data.grid(row=6, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_pswrd = ttk.Label(new_page_window, text="Password", font=("Arial", 12))
        new_pswrd.grid(row=7, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_pswrd_data = ttk.Entry(new_page_window, textvariable=self.new_password_var, width=30, font=("Arial", 12))
        new_pswrd_data.grid(row=7, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_status = ttk.Label(new_page_window, text="Status", font=("Arial", 12))
        new_status.grid(row=8, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        new_status_data = ttk.Combobox(new_page_window, textvariable=self.new_status_var, width=28, font=("Arial", 12), values=('', 'Student', 'Staff'))
        new_status_data.grid(row=8, column=2, columnspan=3, padx=15, pady=5, sticky='w')

        new_register_button = ttk.Button(new_page_window, text="Register new data", command=self.register_data)
        new_register_button.grid(row=9, column=1, columnspan=4, padx=5, pady=20, ipadx=5, ipady=5)

    def register_data(self):
        # Retrieve data from the entry and combobox widgets
        new_id = self.new_id_var.get()
        new_name = self.new_name_var.get()
        new_contact = self.new_contact_var.get()
        new_faculty = self.new_faculty_var.get()
        new_gender = self.new_gender_var.get()
        new_college = self.new_college_var.get()
        new_password = self.new_password_var.get()
        new_status = self.new_status_var.get()

        # Database connection details
        connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\SCSM11\Downloads\Telegram Desktop\Project_Finalize\Project_Finalize\registration database.accdb;'

        try:
            # Connect to the Access database
            conn = pypyodbc.connect(connection_string)
            cursor = conn.cursor()

            # Insert new data into the database
            cursor.execute('INSERT INTO [User] ([UserID], [Name], [Contact], [Faculty], [Gender], [College], [Password], [Status]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)', (new_id, new_name, new_contact, new_faculty, new_gender, new_college, new_password,new_status))

            # Commit the changes to the database
            conn.commit()

            # Display a success message
            messagebox.showinfo("Success", "New data registered successfully!")

        except pypyodbc.Error as e:
            # Display an error message in case of an exception
            messagebox.showerror("Database Error", f"Error registering new data: {e}")

        finally:
            # Close the database connection
            conn.close()

    def edit_data_page(self):
        # Create a new window for the edit data page
        edit_page_window = tk.Toplevel(self.root)
        edit_page_window.title("UiTM Kedah InCoMS | Data Editor Page")

        window_label = ttk.Label(edit_page_window, text="Database Editor", font=("Arial", 14))
        window_label.grid(row=0, column=1, columnspan=4, padx=20, pady=20)

        # Create a listbox to display names and IDs
        listbox = tk.Listbox(edit_page_window, selectmode=tk.SINGLE, font=("Arial", 12), width=60, height=10)
        listbox.grid(row=1, column=1, columnspan=3, padx=15, pady=5, sticky='w')

        # Database connection details
        connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\SCSM11\Downloads\Telegram Desktop\Project_Finalize\Project_Finalize\registration database.accdb;'

        try:
            # Connect to the Access database
            conn = pypyodbc.connect(connection_string)
            cursor = conn.cursor()

            # Query the database to retrieve names and userIDs
            cursor.execute('SELECT [Name], [UserID] FROM [User]')
            user_data = cursor.fetchall()

            # Populate the listbox with names and userIDs
            for name, user_id in user_data:
                listbox.insert(tk.END, f"{user_id} - {name}")

        except pypyodbc.Error as e:
            # Display an error message in case of an exception
            messagebox.showerror("Database Error", f"Error retrieving data: {e}")

        finally:
            # Close the database connection
            conn.close()

        # Button to edit selected data
        edit_button = ttk.Button(edit_page_window, text="Edit Selected Data", command=lambda: self.edit_selected_data(listbox.get(listbox.curselection())))
        edit_button.grid(row=2, column=1, columnspan=3, padx=5, pady=20, ipadx=5, ipady=5)

    def edit_selected_data(self, selected_data):
        # Extract the selected user ID from the string in the listbox
        selected_user_id = selected_data.split(" - ")[0]

        # Query the database to retrieve data for the selected user ID
        connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\SCSM11\Downloads\Telegram Desktop\Project_Finalize\Project_Finalize\registration database.accdb;'
        try:
            conn = pypyodbc.connect(connection_string)
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM [User] WHERE UserID = ?', (selected_user_id,))
            user_data = cursor.fetchone()

            self.open_edit_window(user_data)

        except pypyodbc.Error as e:
            messagebox.showerror("Database Error", f"Error retrieving selected data: {e}")

        finally:
            conn.close()

    def open_edit_window(self, user_data):
        if user_data is None:
            messagebox.showerror("Data Error", "Selected data not found.")
            return
        
        # Create a new window for editing data
        edit_window = tk.Toplevel(self.root)
        edit_window.title("UiTM Kedah InCoMS | Data Editor Page")

        window_label = ttk.Label(edit_window, text="Update selected data", font=("Arial", 14))
        window_label.grid(row=0, column=1, columnspan=4, padx=20, pady=20)

        up_id = ttk.Label(edit_window, text="Student / Staff ID", font=("Arial", 12))
        up_id.grid(row=1, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_id_data = ttk.Entry(edit_window, textvariable=self.up_id_var, width=30, font=("Arial", 12))
        up_id_data.grid(row=1, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_id_data.insert(0, user_data[1])

        up_name = ttk.Label(edit_window, text="Name", font=("Arial", 12))
        up_name.grid(row=2, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_name_data = ttk.Entry(edit_window, textvariable=self.up_name_var, width=30, font=("Arial", 12))
        up_name_data.grid(row=2, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_name_data.insert(0, user_data[4])

        up_contact = ttk.Label(edit_window, text="Contact Number", font=("Arial", 12))
        up_contact.grid(row=3, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_contact_data = ttk.Entry(edit_window, textvariable=self.up_contact_var, width=30, font=("Arial", 12))
        up_contact_data.grid(row=3, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_contact_data.insert(0, user_data[5])

        up_faculty = ttk.Label(edit_window, text="Faculty", font=("Arial", 12))
        up_faculty.grid(row=4, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_faculty_data = ttk.Combobox(edit_window, textvariable=self.up_faculty_var, width=28, font=("Arial", 12), values=('', 'Bussines', 'Engineering', 'Finance', 'Science', 'Pharmacy', 'Mathematics'))
        up_faculty_data.grid(row=4, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_faculty_data.insert(0, user_data[7])

        up_gender = ttk.Label(edit_window, text="Gender", font=("Arial", 12))
        up_gender.grid(row=5, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_gender_data = ttk.Combobox(edit_window, textvariable=self.up_gender_var, width=28, font=("Arial", 12), values=('', 'Male', 'Female'))
        up_gender_data.grid(row=5, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_gender_data.set(user_data[6])

        up_college = ttk.Label(edit_window, text="College", font=("Arial", 12))
        up_college.grid(row=6, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_college_data = ttk.Combobox(edit_window, textvariable=self.up_college_var, width=28, font=("Arial", 12), values=('', 'Kolej Mahsuri', 'Kolej Murni', 'Kolej Masria', 'Kolej Malinja'))
        up_college_data.grid(row=6, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_college_data.set(user_data[8])

        up_pswrd = ttk.Label(edit_window, text="Password", font=("Arial", 12))
        up_pswrd.grid(row=7, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_pswrd_data = ttk.Entry(edit_window, textvariable=self.up_password_var, width=30, font=("Arial", 12))
        up_pswrd_data.grid(row=7, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_pswrd_data.insert(0, user_data[2])

        up_status = ttk.Label(edit_window, text="Status", font=("Arial", 12))
        up_status.grid(row=8, column=1, columnspan=1, padx=15, pady=5, sticky='w')
        up_status_data = ttk.Combobox(edit_window, textvariable=self.up_status_var, width=28, font=("Arial", 12), values=('', 'Student', 'Staff'))
        up_status_data.grid(row=8, column=2, columnspan=3, padx=15, pady=5, sticky='w')
        up_status_data.set(user_data[3])

        update_button = ttk.Button(edit_window, text="Update data", command=self.update_data)
        update_button.grid(row=9, column=1, columnspan=4, padx=5, pady=20, ipadx=5, ipady=5)

    def update_data(self):
        # Retrieve data from the entry and combobox widgets
        user_id = self.up_id_var.get()
        name = self.up_name_var.get()
        contact = self.up_contact_var.get()
        faculty = self.up_faculty_var.get()
        gender = self.up_gender_var.get()
        college = self.up_college_var.get()
        password = self.up_password_var.get()
        status = self.up_status_var.get()

        # Database connection details
        connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\SCSM11\Downloads\Telegram Desktop\Project_Finalize\Project_Finalize\registration database.accdb;'

        try:
            # Connect to the Access database
            conn = pypyodbc.connect(connection_string)
            cursor = conn.cursor()

            # Update data in the database
            cursor.execute('UPDATE [User] SET [Name] = ?, [Contact] = ?, [Faculty] = ?, [Gender] = ?, [College] = ?, [Password] = ?, [Status] = ? WHERE [UserID] = ?', (name, contact, faculty, gender, college, password, status, user_id))

            # Commit the changes to the database
            conn.commit()

            # Display a success message
            messagebox.showinfo("Success", "Data updated successfully!")

        except pypyodbc.Error as e:
            # Display an error message in case of an exception
            messagebox.showerror("Database Error", f"Error updating data: {e}")

        finally:
            # Close the database connection
            conn.close()


if __name__ == "__main__":
    root = tk.Tk()
    app = CollegeManagementSystem(root)
    root.mainloop()