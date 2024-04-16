import os
import shutil
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import process_attendance as pa
    
def browse_file(entry):
    file_selected = filedialog.askopenfilename()
    if file_selected:
        entry.delete(0, tk.END)
        entry.insert(0, file_selected)
    
def browse_folder(entry):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry.delete(0, tk.END)
        entry.insert(0, folder_selected)

def process_data():
    global name_lecturer, tel_no_lecturer, faculty, folder_path, signature_path, exclude_path

    
    # Retrieve data from entry fields
    name_lecturer = name_entry.get()
    tel_no_lecturer = tel_no_entry.get()
    faculty = faculty_entry.get()
    folder_path = folder_entry.get()
    signature_path = signature_entry.get()
    exclude_path = exclude_entry.get()

    # Input validation (you might want to do this)
    if not name_lecturer \
        or not tel_no_lecturer \
        or not faculty \
        or not folder_path \
        or not signature_path \
        or not exclude_path \
            :
        
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    # Close the window
    root.destroy()

    # Start your Processing (replace with your actual actions)
    print("Processing data...") 
    print("Name:", name_lecturer)
    print("Telephone Number:", tel_no_lecturer)
    print("Faculty:", faculty)
    print("Folder Path:", folder_path)
    print("Signature Path:", signature_path)
    print("Exclude Path:", exclude_path)
    
    pass


def on_closing():
    global terminate_program
    terminate_program = 1
    root.destroy()  # Properly terminate the program
    root.quit() 
    

def main():
    global name_entry, tel_no_entry, faculty_entry, folder_entry, signature_entry, exclude_entry
    global root
    
    # GET INPUT FROM USER VIA GUI
    
    # Create the main window
    root = tk.Tk()
    root.title("UTM Attendance Processing Tool")

    # Information label
    info_label = tk.Label(root, 
        text="This program process lecture attendance record...\n"
            + "\n"
            + "Enter your data below:")

    info_label.pack(pady=10)

    # Input fields

    #  Name
    name_label = tk.Label(root, text="Name:")
    name_label.pack()
    name_entry = tk.Entry(root, width=40)
    name_entry.insert(0, "DR. A")
    name_entry.pack()

    #  Telephone Number
    tel_no_label = tk.Label(root, text="Telephone Number:")
    tel_no_label.pack()
    tel_no_entry = tk.Entry(root, width=20)
    tel_no_entry.insert(0, "012-3456789")
    tel_no_entry.pack()

    # Faculty
    faculty_label = tk.Label(root, text="Faculty:")
    faculty_label.pack()
    faculty_entry = tk.Entry(root, width=40)
    faculty_entry.insert(0, "Fakulti Kejuruteraan Mekanikal")
    faculty_entry.pack()

    spacer = tk.Label(root, text="")
    spacer.pack()


    # Folder path with PDFs
    folder_label = tk.Label(root, text="Folder with PDFs:")
    folder_label2 = tk.Label(root, text="Pdf files inside the folder must be named with "
        + "YYMMDD-HH-D format,\n"
        + "e.g. 240331-14-2 corresponds to date 31 March 2024, time 14:00, and 2 hours duration.")
    folder_label2.configure(font=("Helvetica", 9))
    folder_label.pack()
    folder_label2.pack()
    folder_entry = tk.Entry(root, width=60)
    folder_entry.configure(font=("Helvetica", 9))
    folder_entry.insert(0, os.path.join(os.getcwd(), "pdf"))
    folder_entry.pack()

    folder_browse_button = tk.Button(root, text="Browse", 
        command=lambda: browse_folder(folder_entry))
    folder_browse_button.pack()

    spacer = tk.Label(root, text="")
    spacer.pack()


    # Signature path with PDFs
    signature_label = tk.Label(root, text="Signature file (PNG format):")
    signature_label.pack()
    signature_entry = tk.Entry(root, width=70)
    signature_entry.insert(0, os.path.join(os.getcwd(), "signature.png"))
    signature_entry.configure(font=("Helvetica", 9))
    signature_entry.pack()

    signature_browse_button = tk.Button(root, text="Browse", 
        command=lambda: browse_file(signature_entry))
    signature_browse_button.pack()

    spacer = tk.Label(root, text="")
    spacer.pack()
    
    
    # Attendance exclude
    exclude_label = tk.Label(root, text="Attendance exclusion spreadsheet \n(due to MCs or other approved activities):")
    exclude_label.pack()
    exclude_entry = tk.Entry(root, width=70)
    exclude_entry.insert(0, os.path.join(os.getcwd(), "attendance_exclude.xlsx"))
    exclude_entry.configure(font=("Helvetica", 9))
    exclude_entry.pack()

    exclude_browse_button = tk.Button(root, text="Browse", 
        command=lambda: browse_file(exclude_entry))
    exclude_browse_button.pack()

    spacer = tk.Label(root, text="")
    spacer.pack()
    
    
    terminate_program = 0

    root.protocol("WM_DELETE_WINDOW", on_closing)  # Call on_closing when 'X' is clicked


    # Process button
    process_button = tk.Button(root, text="PROCESS", command=process_data)
    process_button.pack(pady=15)


    
    root.mainloop()
    
    if terminate_program:
        return
    
    # MAIN PROCESSING
    
    
    dates = []
    data_dict = {}
    
    output_folder_txt = "txt"
    output_filename = "attendance_processed"
    newest_file = sorted(os.listdir(folder_path))[-1]
    output_filename += '_' + newest_file.strip('.pdf')
    
    
    if os.path.exists(output_folder_txt):
        shutil.rmtree(output_folder_txt)
    if os.path.exists(output_filename+'.csv'):
        os.remove(output_filename+'.csv')
    if os.path.exists(output_filename+'.xlsx'):
        os.remove(output_filename+'.xlsx')
    
    pa.convert_pdfs_to_text(folder_path, output_folder_txt)
    
    pa.extract_data(output_folder_txt, data_dict, dates, exclude_path)

    pa.generate_csv(data_dict, dates, output_filename)

    pa.generate_xlsx(output_filename)

    print(f"{output_filename}.csv and {output_filename}.xlsx are successfully generated.")
    
    
    reminder_letter_folder = 'reminder_letter-generated'
    if os.path.isdir(reminder_letter_folder):
        shutil.rmtree(reminder_letter_folder)    
    os.makedirs(reminder_letter_folder)

    for key, value in data_dict.items():
        
        name_student = key.replace('/', '_')
        path_prefix = f"{reminder_letter_folder}/{name_student}/{value['CourseCode']}-{value['Section']}-{name_student}"
        
        if value['AbsentDuration'] >= int(value['CourseCode'][7])*1:
            os.makedirs(f'{reminder_letter_folder}/{name_student}')
            pa.write_warning_letter(
                name_student, 1, f"{path_prefix}-1st_reminder.pdf", 
                value, name_lecturer, tel_no_lecturer, signature_path)
        
        if value['AbsentDuration'] >= int(value['CourseCode'][7])*2:
            pa.write_warning_letter(
                name_student, 2, f"{path_prefix}-2nd_reminder.pdf", 
                value, name_lecturer, tel_no_lecturer, signature_path)
            
        if value['AbsentDuration'] >= int(value['CourseCode'][7])*3:
            pa.write_warning_letter(
                name_student, 2, f"{path_prefix}-3rd_reminder.pdf", 
                value, name_lecturer, tel_no_lecturer, signature_path)
    
    
    # POST-PROCESSING
    
    messagebox.showinfo("Success", "Attendance records have been processed.")


if __name__ == "__main__":
    main()

