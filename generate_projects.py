#input excel file: projects.xlsx
#template file to be used project-templates.xlsx

import os
import pandas as pd
from openpyxl import load_workbook



if os.path.isdir('output/') is False:
    os.makedirs('output/')

total_projects_saved = 0

projects = pd.read_excel("projects.xlsx")


#print(projects)



workbook = load_workbook(filename="project-template.xlsx")
sheet = workbook.active


#for each project in projects
for ind in projects.index:
    
    sheet["B4"] = projects['ID'][ind]
    sheet["B5"] = projects['NAME'][ind]

    output_filename = "output/" + projects['CONCAT'][ind] + ".xlsx"

    workbook.save(filename=output_filename)
    print("Saved file: " + output_filename)

    total_projects_saved += 1
