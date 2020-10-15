#The Admissions Office Wendy received student information from various regions in China this year, under the path /Users/Wendy/student_information.

#Each student's profile is named after their name, for example: Mao Dapeng.docx

#Now Wendy wants to categorize the student information according to their region so that it can be sent to the person in charge of the corresponding region. The classified folder needs to be created under the path /Users/Wendy/student_information, for example: /Users/Wendy/student information/Sichuan

#The information of which region each student belongs to is in the Excel table /Users/Wendy/student_area.xlsx, the table format is as shown below.

#Please use the knowledge you have learned to help Wendy complete the classification of students. The working directory of the current program is /Users/Wendy/


#Goal: category students' profile files according to their regions.

# import
import openpyxl
import os
import shutil

#set file path
filepath = '/Users/Wendy/student_information'
#read all files in filepath
allfiles = os.listdir(filepath)

#read inforamation workbook
info = openpyxl.load_workbook('/Users/Wendy/student_area.xlsx', data_only =True)
#read worksheet
info_ws = info['Student Area Sheet']

#create a dictionary
name_area = {}

#read the work sheet and find out where are the students come from and put in the dictionary
for item in info_ws.rows:
    student_name = item[0].value
    area = item[1].value
    if student_name != 'name':
        name_area[student_name] = area

#iterate the files in allfiles and match their area with name_area dictionary        
for file in allfiles:
    docpath = os.path.join(filepath, file)
    name = os.path.splitext(file)[0]
    if name in name_area.keys():
        #create the target path which we will move our file to: /Users/Wendy/student information/area
        target_path = os.path.join(filepath, name_area[name])
        #create the directory if it is not exist
        if not os.path.exists(target_path):
            os.mkdir(target_path)
            #move our file to the target_path according to the student's area
            shutil.move(docpath, target_path)
        else:
            #move our file to the target_path according to the student's area
            shutil.move(docpath, target_path)
        
