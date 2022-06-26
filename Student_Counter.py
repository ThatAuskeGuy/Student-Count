import PySimpleGUI as sg
import datetime
import pandas as pd 
import xlsxwriter


sg.theme('Dark Blue 3')

# All the stuff inside your window.
layout = [  
            [sg.Text('Import Debit Account Transactions file', size=(30, 1)), sg.Input(size=(43, 1)), sg.FileBrowse('File Browse', size=(12, 1))],
            [sg.Text('How many Transactions(Activities) to count?'), sg.DropDown(['Choose', '1', '2', '3', '4', '5', '6', '7'], default_value='Choose', size=(10, 1))],
            [sg.Text('Customer: '), sg.Input(size=(10, 1)), sg.Text('Location: '), sg.Input(size=(10, 1)), sg.Text('Invoice Number: '), sg.Input(size=(10, 1)), sg.Text('Cost: '), sg.Input(size=(10, 1))],
            [sg.Button('Start'), sg.Button('Cancel')],
        ]

# Create the Window
window = sg.Window('Student Activity Counter - Version 1.0', layout)
# Event Loop to procepyinstallerss "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    if event == 'Start':
        file = values[0]
        num_or_more = int(values[1])
        customer = values[2]
        location_id = values[3]
        invoice = values[4]
        cost = float(values[5])

        if num_or_more == 'Choose':
            pass
        else:

            """ Open and read the CSV file of data """
            col_names = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF',]
            # D = Start Date, E = Stop Date, I = Name, K = Num of Transactions(Activities)
            df = pd.read_csv(file, names=col_names, keep_default_na=False)
            df["K"] = df["K"].astype(int)
            activity_count = pd.Series(df.K.values, index=df.H).to_dict()
            
            list_of_students = [(key, value) for (key, value) in activity_count.items() if value >= num_or_more]
            dict_of_students = dict(list_of_students)

            start_date = str(df.iloc[0]['C'])
            stop_date = str(df.iloc[0]['D'])

            current_date = datetime.datetime.now()
            if current_date.month == 1:
                invoice_date = f'December {current_date.year - 1}'
            else:
                invoice_date = f"{datetime.date(current_date.year, current_date.month - 1, 1).strftime('%B')} {current_date.year}"


            """ Create Excel sheet to save in customers folder """
            # these two lines generate the workbook/worksheet to write to
            workbook = xlsxwriter.Workbook(f'../2 - Complete - 2/{current_date.year}_{current_date.month}/{invoice} {customer} - {num_or_more} or More Student Activity Count.xlsx')
            workbook.formats[0].set_font_size(10)
            worksheet = workbook.add_worksheet()

            # cell formatting
            page_name_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 14, 'align': 'center'})
            metadata_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'center', 'underline': 1})
            header_cell_formatL = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'left', 'bold': True})
            header_cell_formatR = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'right', 'bold': True})
            student_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'left'})
            activities_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'right'})
            bottom_text_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'center', 'bold': True})

            # report total calculation
            report_total = '${:,.2f}'.format(len(list_of_students) * cost)

            """ Set column widths """
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:B', 40)
            worksheet.set_column('C:C', 20)
            worksheet.set_column('D:D', 6)
            worksheet.set_column('E:E', 15)

            """ Heading, Metadata, and Footer for the sheet and Merge cells """
            worksheet.merge_range('A1:E1', 'ACTIVITIES PER STUDENT', page_name_cell_format)
            worksheet.merge_range('A2:E2', f'CUSTOMER: {customer}', metadata_cell_format)
            worksheet.merge_range('A3:E3', f'LOCATION: {location_id}', metadata_cell_format)
            worksheet.merge_range('A4:E4', f'INVOICE: {invoice}', metadata_cell_format)
            worksheet.merge_range('A5:E5', f'{start_date}    {stop_date}', metadata_cell_format)
            worksheet.merge_range('A6:E6', f'CRITERIA: {num_or_more} or more activities in a month', metadata_cell_format)
            worksheet.write('B8', 'STUDENT NAME', header_cell_formatL)
            worksheet.write('C8', '# STUDENTS:', header_cell_formatR)
            worksheet.write('D8', len(list_of_students), header_cell_formatR)

            """ Add CSV file data to the Excel sheet """
            row = 8
            col = 1

            for student, count in dict_of_students.items():
                worksheet.write(row, col, student, student_cell_format)
                worksheet.write(row, col + 1, '# Activities:', activities_cell_format)
                worksheet.write(row, col + 2, count, activities_cell_format)
                row += 1   
            row += 1     
            
            """ Add report totals to sheets """
            worksheet.merge_range(f'A{row}:E{row}', f'Total Number of Active Students: {len(list_of_students)}', bottom_text_cell_format)
            row += 1 
            worksheet.merge_range(f'A{row}:E{row}', f"Monthly Fee per active student: {'${:,.2f}'.format(cost)}", bottom_text_cell_format)
            row += 1 
            worksheet.merge_range(f'A{row}:E{row}', f'Fee for {invoice_date}: {report_total}', bottom_text_cell_format)
            row += 1 

            workbook.close()