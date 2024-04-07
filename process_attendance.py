"""
Copyright (c) 2024, Hazmil Azis, hazmil.abdazis@gmail.com

Licensed under the MIT License.
For more information, see the LICENSE.txt file.

Script to extract attendance information from PDF files generated from UTM 
attendance recording system and generate an Excel file compiling daily attendance.

The script converts PDF files in the folder to text, extracts the relevant 
data, and saves the data to an Excel file.
The Excel file will have one sheet with the converted data.
The script assumes the PDF files are named in the format "<date>.pdf" 
where <date> is in the format YYMMDD.
The first row of the Excel file will have the header row with the column names.
"""



import os
import shutil
import subprocess
import csv
import xlsxwriter
import pdftotext


def convert_pdfs_to_text(input_folder, output_folder):
    """Converts PDFs in a folder to text files.

    Args:
        input_folder (str): Path to the folder containing PDF files.
        output_folder (str): Path to the folder where text files will be saved.
    """


    os.makedirs(output_folder, exist_ok=True)  # Create the output folder if it doesn't exist

    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(input_folder, filename)

            output_filename = os.path.splitext(filename)[0] + ".txt"
            output_path = os.path.join(output_folder, output_filename)

            if 0: # to replace
                command = ["pdftotext", "-layout", pdf_path, output_path]
                result = subprocess.run(command, capture_output=True, text=True)
                if result.returncode != 0:
                    print(f"Error converting PDF: {result.stderr}")
                else:
                    print(f"PDF converted to text: {output_path}")
            elif 1: # to replace
                with open(pdf_path, 'rb') as pdf_file, open(output_path, 'w') as text_file:
                    pdf = pdftotext.PDF(pdf_file, physical=True)
                    text = "\n\n".join(pdf)
                    text_file.write(text)


            print(f"Converted {pdf_path} to {output_path}")
            
            
def extract_data(input_folder, data_dict: dict, dates: list):
    """Extract data from text files in a folder.

    Args:
        input_folder (str): Path to the folder containing text files.
        data_dict (dict): A dictionary to store data.
        dates (list): A list to store dates.
    """
    
    # A loop that will open all files inside a folder
    listfiles = os.listdir(input_folder)
    listfiles.sort(reverse=True)
    
    latest_date_done = False
    
    for filename in listfiles:
        if filename.lower().endswith('.txt'):
            file_in_path = os.path.join(input_folder, filename)
            
            print("Processing " + file_in_path + " ...")
        
            with open(file_in_path, 'r') as infile, \
                open(file_in_path+'.stripped', 'w') as outfile:
                
                
                lines = infile.readlines()
                
                # remove empty lines
                lines = [line for line in lines if line.strip()]
                outfile.writelines(lines)
                
                # Assign content of lines from row 0 to 5
                header = lines[0:5]
                
                date_time_split = lines[4].split()
                date_time = " ".join(date_time_split[-5:])
                dates.append(date_time)
                
                header_table = lines[6]
                
                
                # Assign content of lines from row 11 until the end of the table, 
                #   but exclude the two last line
                data = lines[7:-1]
                data_size = len(data)
                
                
                for line in data:
                    
                    if any(c.strip() for c in line[:4]):
                        # This is data row
                        
                        line_words = line.split()
                        no = line_words[0]
                        matric_no = line_words[1]
                        
                        if line_words[-1][-1] == "M":
                            # row containing "AM" or "PM" means attended
                            time_in = ' '.join(line_words[-2:])
                            year = line_words[-3]
                            programme = line_words[-4]
                            name = ' '.join(line_words[2:-4])
                            
                        else:
                            time_in = ""
                            year = line_words[-1]
                            programme = line_words[-2]
                            name = ' '.join(line_words[2:-2])
                            
                    if not latest_date_done:
                        if name not in data_dict:
                            data_dict[name] = {'Name': name}
                            data_dict[name]['MatricNo.'] = matric_no
                            data_dict[name]['Programme'] = programme
                            data_dict[name]['Year'] = year
                            data_dict[name]['Attended'] = 0
                            data_dict[name]['Absent'] = 0
                            data_dict[name]['AbsentList'] = ""
                    
                    if name not in data_dict:
                        print('Name is not in the latest name list:', name)
                        print('This row will be ignored:', line)
                        
                    else:
                        data_dict[name][date_time] = time_in
                    
                        if time_in != '':
                            data_dict[name]['Attended'] += 1
                        else:
                            data_dict[name]['Absent'] += 1
                            data_dict[name]['AbsentList'] += date_time + '; '
                        
                if not latest_date_done:
                    latest_date_done = True
                        
        pass
                        
                        
def generate_csv(data_dict: dict, dates: list, output_filename: str):
    """
    Generate a csv file based on the data extracted from text files.

    Args:
        data_dict (dict): A dictionary of students data.
        dates (list): A list of dates.
        output_filename (str): The name of the output csv file.

    """

    
    with open(output_filename+'.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        
        fieldnames = \
            ['No.'] \
            + ['Name', 'MatricNo.', 'Programme', 'Year'] \
            + ['Attended', 'Absent', 'Percentage', 'AbsentList'] \
            + dates

        writer.writerow(fieldnames)
        
        row_count = 0
        
        for key, value in data_dict.items():
            row_count += 1
            
            row_content = [
                    row_count,
                    value['Name'], value['MatricNo.'], value['Programme'], 
                    value['Year']
                ] \
                + [ value['Attended'], value['Absent'] ] \
                + [ "{:.1f}".format(value['Attended']/(value['Attended'] + value['Absent'])*100) ] \
                + [ value['AbsentList'][:-2] ] \
                + [value.get(date, '') for date in dates]
            writer.writerow(row_content)
            
            
            
            

def generate_xlsx(output_filename: str):
    """
    Generate an Excel file based on the data extracted from csv files.

    Args:
        output_filename (str): The name of the output Excel file.

    """
    
    data_csv = []
    with open(output_filename+'.csv', 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            data_csv.append(row)


    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook(output_filename+'.xlsx')
    worksheet = workbook.add_worksheet()

    # Write the list data into the worksheet 
    for row_num, row_data in enumerate(data_csv):
        for col_num, value in enumerate(row_data):
            worksheet.write(row_num, col_num, value)


    # Close the workbook
    workbook.close()


# -------- Main Execution ---------

dates = []
data_dict = {}

input_folder = "pdf"
output_folder_txt = "txt"
output_filename = "attendance_processed"

if os.path.exists(output_folder_txt):
    shutil.rmtree(output_folder_txt)
if os.path.exists(output_filename+'.csv'):
    os.remove(output_filename+'.csv')
if os.path.exists(output_filename+'.xlsx'):
    os.remove(output_filename+'.xlsx')

convert_pdfs_to_text(input_folder, output_folder_txt)

extract_data(output_folder_txt, data_dict, dates)

generate_csv(data_dict, dates, output_filename)

generate_xlsx(output_filename)

print(f"{output_filename}.csv and {output_filename}.xlsx are successfully generated.")

def display_attendance_data(filename):
    import tkinter as tk
    from tkinter import ttk  # For the Treeview widget
    from tkinter import font
    import csv

    def read_csv_data(filename):
        data = []
        with open(filename, 'r') as file:
            reader = csv.reader(file)
            header = next(reader)  # Read the header row
            data = [row for row in reader]  # Read the data rows
        return header, data

    def create_table(root, header, data):

        tree = ttk.Treeview(root, columns=header, show='headings')
        
        # Define headings
        for col in header:
            tree.heading(col, text=col)  

        # Add data rows
        for row in data:
            tree.insert('', tk.END, values=row)

        tree.pack() 
        

    # Main Tkinter window
    root = tk.Tk()
    root.title("CSV Table Viewer")
    
    # Better default font 
    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(family="Arial", size=9)

    # Example usage
    csv_file = filename
    header, data = read_csv_data(csv_file)
    create_table(root, header, data)

    root.mainloop()
    

display_attendance_data(output_filename+'.csv')