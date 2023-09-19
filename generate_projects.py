#input excel file: projects.xlsx
#template file to be used project-templates.xlsx

import os
import pandas as pd
from openpyxl import load_workbook


template_excel = "In_Service_Form.xlsx"
project_list = "Sept_In_Service_To_Dos.xlsx"


if os.path.isdir('output/') is False:
    os.makedirs('output/')

total_projects_saved = 0

projects = pd.read_excel(project_list)


#print(projects)



workbook = load_workbook(filename=template_excel)
sheet = workbook.active


#for each project in projects
for ind in projects.index:
    if type(projects["Concatonate"][ind]) == float:
        continue
    if len(projects["Concatonate"][ind]) <= 1:
        continue


    sheet["D4"] = projects.iloc[ind,0]
    sheet["D5"] = projects.iloc[ind,1]
    sheet["B12"] = projects.iloc[ind,2]

    output_filename = "output/" + projects.iloc[ind,2] + ".xlsx"

    workbook.save(filename=output_filename)
    print("Saved file: " + output_filename)

    total_projects_saved += 1

    
