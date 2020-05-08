"""               AUTO CREDIT
        Excel to PDF credit list generator
      Written and maintained by Thomas Philip
                    -lockdown project 2020
"""
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import datetime

# create a new Excel workbook to store the results
result_workbook = openpyxl.Workbook()
# create a new excel sheet inside result_workbook
result_sheet = result_workbook.active   # the result sheet stores the result data

# in the result sheet, merge cells from A1 to H1 (row 1, col 1-8)
result_sheet.merge_cells('A1:H1') # (i.e. a rectangular field is created) for the main title
# find the current date for title
current_date = datetime.datetime.today().date()
# main title for the result sheet >>>  "PENDING PAYMENTS : dd-mm-yyyy"
result_sheet.cell(row = 1, column = 1).value = 'PENDING PAYMENTS : '+ current_date.strftime('%d-%m-%Y') #(title is in the 1st row)

# give the column names for the result sheet   (column names are in the 2nd row)
# Date, Inv No, Retailor Name, Address, Amount, Disc., Received, PENDING  :(total 8 columns)
result_sheet.cell(row = 2, column = 1).value="Date"
result_sheet.cell(row = 2, column = 2).value="Inv No."
result_sheet.cell(row = 2, column = 3).value="Retailor Name"
result_sheet.cell(row = 2, column = 4).value="Address"
result_sheet.cell(row = 2, column = 5).value="Amount"
result_sheet.cell(row = 2, column = 6).value="Disc."
result_sheet.cell(row = 2, column = 7).value="Received"
result_sheet.cell(row = 2, column = 8).value="PENDING"

# Get the location of the source excel file
source_excel_path = "C:\\Users\\thoma\\OneDrive\\Desktop\\excel python\\test.xlsx"
# double '\\' for path is used, since single '\' is used for escape sequencing

# load the source excel workbook from the given source path
sorce_workbook = openpyxl.load_workbook(source_excel_path)
# store all the sheet names of the source workbook to a list (source_sheetnames)
source_sheetnames=sorce_workbook.sheetnames # this list contains sheetnames A to Z

# rn changed to result_current_row
result_current_row = 3  # inside the result sheet, the rows used for storing the cell values starts from 3,
# since rows 1 & 2 are used for main title & column headings

""" this code below, has to be looked upon later """
# for amt and rmax code, it is to know from which row there is no str fields in received column of each sheet. eg: 23450+3200 is str field
rlst=[290,320,21,25,89,87,91,2,2,297,355,2,290,67,2,230,2,65,360,123,2,2,2,2,2,2] # 26 values for each sheet A to Z
i=0 # for indexing rlst
""" ............................................. """

# (1st) for loop to go through each sheet inside the source workbook
for character in source_sheetnames:
    source_sheet = sorce_workbook[character]  # sheets A to Z in source workbook is assigned to source_sheet after each iteration

    source_max_row = source_sheet.max_row # gets the maximum no. of rows in the current source_sheet

    # (2nd) for loop to iterate through each row from '2' to 'source_max_row' in the current source_sheet
    for source_current_row in range(rlst[i], source_max_row+1):  #source_current_row stores the no. of the current row
        """ above the range starts from rlst[i] instead of 2, to ensure that there is no strings or invalid type values in received column """
        """ try for a better approach for the above problem, try cleaning the source excel file to make sure the received column contains a single integer value"""

        # store the cell(column) date of the current row
        source_cell_date = source_sheet.cell(row = source_current_row, column = 1).value # date
        # check if date is not None and if so, store all the column values of the current row
        if source_cell_date != None :
            source_cell_inv_no = source_sheet.cell(row = source_current_row, column = 2).value # inv.no
            source_cell_retailor_name = source_sheet.cell(row = source_current_row, column = 3).value # Retailor name
            source_cell_address = source_sheet.cell(row = source_current_row, column = 4).value # Address
            source_cell_amount = source_sheet.cell(row = source_current_row, column = 5).value # Amount
            source_cell_discount = source_sheet.cell(row = source_current_row, column = 6).value # Discount
            # check if received column is None or not
            if  source_sheet.cell(row = source_current_row, column = 7).value == None:
                source_cell_received = 0 # if received column is None, assign received = 0
            else:
                source_cell_received = source_sheet.cell(row = source_current_row, column = 7).value # if not None, assign received it's column value
            # calculate the pending amount (amount - received)
            source_cell_pending = source_cell_amount - source_cell_received

            # to check whether pending amount is greater than 0(or any other min. amount (>0) we want)
            # only if it is True, we move all the stored column values of the current row in the source sheet to the result sheet
            if source_cell_pending > 0 :
                # write column values to result sheet
                result_sheet.cell(row = result_current_row, column = 1).value = source_cell_date.strftime('%d-%m-%Y') # change the Date to string format
                # the date in the string format(if we store date in non-string format, it would result in gibberish) is moved to date column in the result sheet
                result_sheet.cell(row = result_current_row, column = 2).value = source_cell_inv_no # Inv. no
                result_sheet.cell(row = result_current_row, column = 3).value = source_cell_retailor_name # Retailor name
                result_sheet.cell(row = result_current_row, column = 4).value = source_cell_address # Address
                result_sheet.cell(row = result_current_row, column = 5).value = source_cell_amount # Amount
                result_sheet.cell(row = result_current_row, column = 6).value = source_cell_discount # Discount
                result_sheet.cell(row = result_current_row, column = 7).value = source_cell_received # Received
                result_sheet.cell(row = result_current_row, column = 8).value = source_cell_pending # PENDING

                # after writing all the columns in result sheet, increment the result_current_row for moving to the next row in result sheet
                result_current_row += 1


            """ above loops and if statements are properly maintained """

    i += 1       # for rlst , might need to change this

# print statement just for feedback if the complete extraction and writing process was successfully finished
print("   SUCCESS ! ")

""" below code is for formatting and styling the result sheet """

# set the height of the 1st row for main title
result_sheet.row_dimensions[1].height = 30
# set the height of the 2nd row for column headings
result_sheet.row_dimensions[2].height = 50

# set the width of each column ('A' means 1st column, 'B' 2nd column and so on...)
result_sheet.column_dimensions['A'].width = 11  # Date column
result_sheet.column_dimensions['B'].width = 11  # Inv. no column
result_sheet.column_dimensions['C'].width = 40  # Retailor name column
result_sheet.column_dimensions['D'].width = 30  # Address column
result_sheet.column_dimensions['E'].width = 14  # Amount column
result_sheet.column_dimensions['F'].width = 13  # Discount column
result_sheet.column_dimensions['G'].width = 14  # Received column
result_sheet.column_dimensions['H'].width = 14  # PENDING column

# set the main title font style to bold
result_sheet.cell(row = 1, column = 1).font = Font(size = 22, bold = True)

#font size for each coloumn
result_sheet.cell(row = 2, column = 1).font = Font(size = 18) # Date column
result_sheet.cell(row = 2, column = 2).font = Font(size = 18) # Inv. no column
result_sheet.cell(row = 2, column = 3).font = Font(size = 18) # Retailor name column
result_sheet.cell(row = 2, column = 4).font = Font(size = 18) # Address column
result_sheet.cell(row = 2, column = 5).font = Font(size = 18) # Amount column
result_sheet.cell(row = 2, column = 6).font = Font(size = 18) # Discount column
result_sheet.cell(row = 2, column = 7).font = Font(size = 18) # Received column
result_sheet.cell(row = 2, column = 8).font = Font(size = 18, bold = True, color='FF0000') # PENDING column, color is changed to red and font is BOLD

# set centre alignment for main title
title_alignment = Alignment(horizontal='center',vertical='bottom',text_rotation=0,wrap_text=False,shrink_to_fit=True,indent=0)
result_sheet.cell(row = 1, column = 1).alignment = title_alignment

""" Save the result workbook to result excel path """
result_excel_path = "C:\\Users\\thoma\\OneDrive\\Desktop\\excel python\\result_excel.xlsx"
result_workbook.save(result_excel_path)
