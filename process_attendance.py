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
import fitz


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
    
    import pandas as pd
    
    # A loop that will open all files inside a folder
    listfiles = os.listdir(input_folder)
    listfiles.sort(reverse=True)
    
    latest_date_done = False
    
    if os.path.isfile('attendance_exclude.xlsx'):
        with open('attendance_exclude.xlsx', 'rb') as file:
            data_exclude = pd.read_excel(file)
    
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
                course_code_name = header[2].strip().replace("\xad", "").split()
                course_code = course_code_name[2]
                course_name = " ".join(course_code_name[3:])
                
                section = header[3].strip().split()[2]

                
                

                date_time = filename[0:-4]
                dates.append(date_time)
                
                duration = int(filename[10])
                
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
                            data_dict[name]['CourseCode'] = course_code
                            data_dict[name]['CourseName'] = course_name
                            data_dict[name]['Section'] = section
                            data_dict[name]['Attendance'] = {}
                            data_dict[name]['AttendanceExcluded'] = []
                            data_dict[name]['Attended'] = 0
                            data_dict[name]['Absent'] = 0
                            data_dict[name]['AbsentList'] = ""
                            data_dict[name]['AbsentDuration'] = 0
                            
                            if os.path.isfile('attendance_exclude.xlsx'):
                                
                                data_exclude_filtered = data_exclude[data_exclude['Name'] == name]
                                exclude_list = data_exclude_filtered['Exclude'].to_string().split()
                                result = []
                                for word in exclude_list:
                                    result += word.split(",")
                                exclude_list = result
                                exclude_list = [x for x in exclude_list if x != '']
                                exclude_list = exclude_list[1:]
                                
                                print("data_exclude_filtered: ")
                                print(data_exclude_filtered)
                                print("exclude_list: ")
                                print(exclude_list)
                                
                                data_dict[name]['AttendanceExcluded'] = exclude_list.copy()
                    
                    
                    if name not in data_dict:
                        print('Name is not in the latest name list:', name)
                        print('This row will be ignored:', line)
                        
                    else:
                        data_dict[name]['Attendance'][date_time] = time_in
                        
                        if time_in == '' and date_time in data_dict[name]['AttendanceExcluded']:
                            time_in = 'Excluded'
                            data_dict[name]['Attendance'][date_time] = time_in
                    
                        if time_in != '':
                            data_dict[name]['Attended'] += 1
                        else:
                            data_dict[name]['Absent'] += 1
                            data_dict[name]['AbsentList'] += date_time + '; '
                            data_dict[name]['AbsentDuration'] += duration
                        
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
            + ['Attended', 'Absent', 'Percentage', 'AbsentList (YYMMDD-HH-Duration)', 'AbsentDuration (Hrs)'] \
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
                + [ value['AbsentDuration'] ] \
                + [value['Attendance'].get(date, '') for date in dates]
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

    
    
def write_warning_letter(
    name_student: str, warning_level: int, write_path: str, value_dict: dict, name_lecturer: str, phone_number: str):
    

    def draw_grid(page):
        r_grid = []; a_grid = []
        w = 25; h = 6
        i = 0
        for x in range(100,501,100):
            for y in range(100,801,100):
                r_grid.append(fitz.Rect(x,y,x+w,y+h))
                a_grid.append(page.add_freetext_annot(
                    r_grid[i], f"{x},{y}", fontsize=6, fill_color=gold))
                i += 1


    is_draw_grid = False

    # some colors
    blue  = (0,0,1)
    green = (0,1,0)
    red   = (1,0,0)
    gold  = (1,1,0)


    if warning_level == 1:
        doc = fitz.open(
            "forms/Peringatan Pertama - Surat Tidak Hadir Kuliah.pdf")
    elif warning_level == 2:
        doc = fitz.open(
            "forms/Peringatan Kedua - Surat Tidak Hadir Kuliah.pdf")
    elif warning_level == 3:
        doc = fitz.open(
            "forms/Peringatan Akhir - Surat Tidak Hadir Kuliah.pdf")
    


    # add annotation to an existing PDF file
    page = doc[0] # get the first page
    page2 = doc[1] # get the second page

    # PAGE 1

    if 0: # example. TO REMOVE
        # the text, Latin alphabet
        t = "¡Un pequeño texto para practicar!"
        # 3 rectangles, same size, above each other
        r1 = fitz.Rect(100,100,200,150)
        r2 = r1 + (0,75,0,75)
        r3 = r2 + (0,75,0,75)
        
        # add 3 annots, modify the last one somewhat
        a1 = page.add_freetext_annot(r1, t)
        a2 = page.add_freetext_annot(r2, t, fontname="Ti")
        a3 = page.add_freetext_annot(r3, t, fontname="Co", rotate=90)
        a3.set_border(width=0)
        a3.update(fontsize=8, fill_color=gold)

    # grid
    if is_draw_grid:
        draw_grid(page)

    # Data to fill in
    
    from datetime import date
    today = date.today()
    today = today.strftime("%d/%m/%Y")
    

    # Date
    x = 430; y = 140; w = 100; h = 20;
    r4 = fitz.Rect(x,y,x+w,y+h)
    a4 = page.add_freetext_annot(r4, today)

    # Nama pelajar
    x = 200; y = 190; w = 300; h = 20;
    r4 = fitz.Rect(x,y,x+w,y+h)
    a4 = page.add_freetext_annot(r4, name_student)

    # No kad matrik
    x = 200; y = 225; w = 200; h = 20;
    r5 = fitz.Rect(x,y,x+w,y+h)
    a5 = page.add_freetext_annot(r5, value_dict['MatricNo.'])

    # Tahun program
    x = 200; y = 265; w = 200; h = 20;
    r6 = fitz.Rect(x,y,x+w,y+h)
    a6 = page.add_freetext_annot(r6, value_dict['Year'])

    # Fakulti
    x = 200; y = 300; w = 200; h = 20;
    r7 = fitz.Rect(x,y,x+w,y+h)
    a7 = page.add_freetext_annot(r7, "Fakulti Kejuruteraan Mekanikal")
    
    
    # Table data
    kod_kursus_list = []
    nama_kursus_list = []
    tarikh_list = []
    jam_list = []
    
    for key, value in value_dict['Attendance'].items():
        if value == "":
            kod_kursus_list.append(value_dict['CourseCode'])
            nama_kursus_list.append(value_dict['CourseName'])
            tarikh_list.append(f"{key[4:6]}/{key[2:4]}/{key[0:2]}")
            jam_list.append(key[10])


    # Kod kursus boxes

    r_kod = []; a_kod = []
    x = 80; w = 70; h = 20
    y = [500, 523, 546, 569, 592, 615]

    for i in range(len(kod_kursus_list)):
        r_kod.append(fitz.Rect(x,y[i],x+w,y[i]+h))
        a_kod.append(page.add_freetext_annot(
            r_kod[i], kod_kursus_list[i], fontsize=10))
        
    # Nama kursus boxes
    r_nama = []; a_nama = []
    x = 160; w = 155; h = 20
    y = [500, 523, 546, 569, 592, 615]
    y = [yi-2 for yi in y]

    for i in range(len(kod_kursus_list)):
        r_nama.append(fitz.Rect(x,y[i],x+w,y[i]+h))
        a_nama.append(page.add_freetext_annot(
            r_nama[i], nama_kursus_list[i], fontsize=9))
        
    # Tarikh boxes
    r_tarikh = []; a_tarikh = []
    x = 325; w = 130; h = 20
    y = [500, 523, 546, 569, 592, 615]
    y = [yi-2 for yi in y]

    for i in range(len(kod_kursus_list)):
        r_tarikh.append(fitz.Rect(x,y[i],x+w,y[i]+h))
        a_tarikh.append(page.add_freetext_annot(
            r_tarikh[i], tarikh_list[i], fontsize=9))
        
    # Jam boxes
    r_jam = []; a_jam = []
    x = 450; w = 130; h = 20
    y = [500, 523, 546, 569, 592, 615]
    y = [yi-2 for yi in y]

    for i in range(len(kod_kursus_list)):
        r_jam.append(fitz.Rect(x,y[i],x+w,y[i]+h))
        a_jam.append(page.add_freetext_annot(
            r_jam[i], jam_list[i], fontsize=10))
        
        
        
    # PAGE 2

    if is_draw_grid:
        draw_grid(page2)


    # Lecturer's name
    x = 170; y = 695; w = 300; h = 20;
    r_lecname = fitz.Rect(x,y,x+w,y+h)
    a_lecname = page2.add_freetext_annot(r_lecname, name_lecturer)


    # Lecturer's phone number
    x = 170; y = 730; w = 300; h = 20;
    r_lectel = fitz.Rect(x,y,x+w,y+h)
    a_lectel = page2.add_freetext_annot(r_lectel, phone_number)
    
    
    # Lecturer's signature
    if os.path.isfile('signature.png'):
        image_path = 'signature.png'
        image_rectangle = fitz.Rect(100, 635, 200, 673)  # (x0, y0, x1, y1)
        page2.insert_image(image_rectangle, filename=image_path)
        


    # save the PDF
    doc.save(write_path)
    
    print(f"Warning letter generated: {write_path}")



# -------- Main Execution ---------
def main():
    dates = []
    data_dict = {}

    name_lecturer = 'DR. MOHD HAZMIL SYAHIDY BIN ABDOL AZIS'
    tel_no_lecturer = '013-7034072'
    
    print(f"name={name_lecturer}")
    print(f"phone_number={tel_no_lecturer}")
    

    input_folder = "pdf"
    output_folder_txt = "txt"
    output_filename = "attendance_processed"
    newest_file = sorted(os.listdir(input_folder))[-1]
    output_filename += '_' + newest_file.strip('.pdf')

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

    if os.path.isdir('reminder_letter'):
        shutil.rmtree('reminder_letter')    
    os.makedirs('reminder_letter')

    for key, value in data_dict.items():
        
        name_student = key.replace('/', '_')
        path_prefix = f"reminder_letter/{name_student}/{value['CourseCode']}-{value['Section']}-{name_student}"
        
        if value['AbsentDuration'] >= 3:
            os.makedirs(f'reminder_letter/{name_student}')
            write_warning_letter(
                name_student, 1, f"{path_prefix}-1st_reminder.pdf", value, name_lecturer, tel_no_lecturer)
        
        if value['AbsentDuration'] >= 6:
            write_warning_letter(
                name_student, 2, f"{path_prefix}-2nd_reminder.pdf", value, name_lecturer, tel_no_lecturer)
            
        if value['AbsentDuration'] >= 9:
            write_warning_letter(
                name_student, 2, f"{path_prefix}-3rd_reminder.pdf", value, name_lecturer, tel_no_lecturer)
            
if __name__ == "__main__":
    main()
