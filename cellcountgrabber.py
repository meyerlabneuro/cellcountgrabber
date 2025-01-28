#BEFORE RUNNING THIS - make an empty CSV file in the csvs folder and name it appropriately 
import os
import csv 
from openpyxl import load_workbook

#direct script to the folder where the excels are 
folder_path = r'C:\Users\Gabrielle\Documents\GitHub\CellCountGrabber\cell_excels\xlsx_pl_cfos_redo'
#extract relevant information from excels
def extract_count_excel(folder_path): #doublecheck your folder path is correct
    data = [] #making empty list to store extracted data

    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            workbook = load_workbook(file_path)
            sheet = workbook.active
            value = sheet['B2'].value
            data.append({'Filename': filename, 'Value_B2': value})
            workbook.close()
    return data   

#store the results into a function
data = extract_count_excel(folder_path)

#path to the empty CSV FILE where you want to put the cell counts
csv_path = r'C:\Users\Gabrielle\Documents\GitHub\CellCountGrabber\csvs\il_cfos_final.csv'

#ensure the directory for the CSV file exists
output_dir = os.path.dirname(csv_path)
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

#put the extract data into csv
with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = ['Filename', 'Value_B2']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=',')
    writer.writeheader()
    for row in data:
        writer.writerow(row)

#have the program tell you where the CSV file ended up 
print(f"CSV file saved to: {csv_path}")