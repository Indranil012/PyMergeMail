#Author : Indranil012
#Email : indranilchowdhury0@gmail.com

from colorama import Fore, Style
import pandas as pd
import os, json, sys, smtplib
from email.message import EmailMessage
from docxtpl import DocxTemplate
import pkg_resources

#check if required modules are installed
required = {"pandas", "docxtpl", "colorama"}
installed = {pkg.key for pkg in pkg_resources.working_set}
missings = required - installed
missing_str = ""
for missing in missings:
    missing_str += " " + missing
if missings:
    print(f"Some package is missing\nPlease install missing packages by\npip install{missing_str}")
    sys.exit()

#Enter the complete filepath to the Word Template
#template_path = input("Insert the Word template file path : ")
template_path = os.getcwd() + "/sampledoc.docx"
print(Fore.GREEN + "\nReading template from: " + template_path)

#To print variables
tpl = DocxTemplate(template_path)
variables = list(tpl.get_undeclared_template_variables())
print(f"\nPrinting variables...\n{variables}")

#Enter the complete filepath of the excel file which has the data
#source_path = input("Insert the source Excel data file path : ")
source_path = os.getcwd() + "/source_data.xlsx"
print("\nReading source data from: " + source_path)
df = pd.read_excel(source_path)
print(f"\nPrinting source data...\n{df}{Style.RESET_ALL}")

#Enter the complete directory path where exported files to be saved
#save_dir = input("Insert the export directory path : ")
save_dir = os.getcwd() + '/export/'

count = 0
for index, row in df.iterrows():
    all_data = row.to_dict()
    context = {}
    for variable in variables:
        if variable in all_data:
            context[variable]=all_data[variable]
    tpl.render(context)
    tpl.save(save_dir + '{}_{}.docx'.format(context["Name"],context["Age"]))
    count += 1
print(f"\n{count} docxs are saved to: {save_dir}")
