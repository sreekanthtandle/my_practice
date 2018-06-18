import xlsxwriter
import pandas as pd
import os



workbook = xlsxwriter.Workbook('strings.xlsx')
worksheet = workbook.add_worksheet()
worksheet.insert_image('B2', '/home/sreekanth/Downloads/sooktha_logo2.png')


class Some_class:
    def __init__(self):


# Sooktha Logo
#   def logo(self):
#         workbook = xlsxwriter.Workbook('strings.xlsx')
#         worksheet = workbook.add_worksheet()       
#         worksheet.insert_image('B2', '/home/sreekanth/Downloads/sooktha_logo2.png')
          
# Closing excel and placing logo.
    def closing(self):
           #workbook = xlsxwriter.Workbook('strings.xlsx')
           #worksheet = workbook.add_worksheet()
           worksheet.insert_image('B2', '/home/sreekanth/Downloads/sooktha_logo2.png')
          
            
# User perspective questions and storage.
    def enter_data(self):
             emp_name =   raw_input("Enter employee name?\r\n")
             todays_date= raw_input("Enter Today's Date?\r\n")
             place =      raw_input("Enter Place?\r\n")
             purpose=     raw_input("Enter purpose of carrying equipment\r\n")
             eqp_return_date = raw_input("Enter equipment return date ?\r\n")

# Location of the excel file.
    def destination_file(self):
             workbook = xlsxwriter.Workbook('strings.xlsx')
             worksheet = workbook.add_worksheet()
                #red = workbook.add_format({'color': 'red'})
                #blue = workbook.add_format({'color': 'blue'})
                #text_wrap = workbook.add_format({'text_wrap': False})
                #string_parts1 = ['To whomsoever it may concern']
                #worksheet.write_rich_string('H7', *string_parts1)

# Address of the office.
    def address(self):
            Sooktha_Address_Line1   = ['Sooktha Consulting Private Limited'] 
            Sooktha_Address_Line1_1 = ['Phone: +91- 080-41114401']
            Sooktha_Address_Line2   = ['57/53/1,2, SV Arcade,'] 
            Sooktha_Address_Line2_2 = ['Fax: +91- 80-41114401']
            Sooktha_Address_Line3   = ['Bilekahalli Main Road, Off Bannerghatta Road,']
            Sooktha_Address_Line4   = ['Bangalore - 560076, Karnataka, INDIA']
                       
 
                  
            cell_format = workbook.add_format()
            cell_format.set_font_name('Times New Roman')
            worksheet.write_rich_string('B5', *Sooktha_Address_Line1)
            worksheet.write_rich_string('H5', *Sooktha_Address_Line1_1)
            worksheet.write_rich_string('B6', *Sooktha_Address_Line2)
            worksheet.write_rich_string('B7', *Sooktha_Address_Line2_2)
            worksheet.write_rich_string('B8', *Sooktha_Address_Line3)
            worksheet.write_rich_string('B9', *Sooktha_Address_Line4)
            


            cell_format.set_font_size(15)
                        
             

# Under Test
'''cell_format = workbook.add_format()
cell_format.set_font_name('Times New Roman')
worksheet.write_rich_string('E4', *Sooktha_Address_Line4)
#worksheet.set_column('E5:N4', 9, cell_format)
cell_format.set_font_size(15)
cell_format.set_underline()
#worksheet.write_rich_string('E3', *string_parts2)
string_parts2 = ['Subject: Authorization to carry ', emp_name[:26], ' equipment for demonstration purposes at Mobile World Congress' ]
string_parts2 = ['Subject: Authorization to carry ', emp_name[:26], ' equipment for demonstration purposes at Mobile World Congress' ]
string_parts2 = ['Subject: Authorization to carry ', emp_name[:26], ' equipment for demonstration purposes at Mobile World Congress' ]
string_parts2 = ['Subject: Authorization to carry ', emp_name[:26], ' equipment for demonstration purposes at Mobile World Congress' ]

cell_format = workbook.add_format()
cell_format.set_font_name('Times New Roman')
#worksheet.set_row(0, 18, cell_format)
worksheet.set_column('H:N', 9, cell_format)
cell_format.set_font_size(15)
cell_format.set_underline()
#cell_format.set_border()
cell_format.set_bold()      # Turns bold on.
cell_format.set_bold(True)  # Also turns bold on.

#worksheet.write_string(2, 3, 'Bar')

#cell_format.set_border()

#worksheet.insert_image('B2', '/home/sreekanth/Downloads/ec200.png')


worksheet.write_rich_string('E9', *string_parts2)

#string_parts = [ 'This is ', emp_name[:16]  , red, 'red', ' and this is ', blue, 'blue']
#worksheet.write_rich_string('C15', *string_parts)
#print (emp_name[:16])

#string_parts.append(text_wrap)
#worksheet.write_rich_string('A2', *string_parts)
#worksheet.write_rich_string('C16',*string_parts)'''

# List of functions:
obj1 = Some_class()

obj1.address()
#obj1.logo()
obj1.enter_data()
obj1.closing()


#workbook.close()


