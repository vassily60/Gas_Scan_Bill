import pdfplumber
import pandas as pd
import calendar
import argparse
import os
import openpyxl


data = {
    'Amount': [],
    'Meter': [],
    'Usage': [123],
}

def parse_pdf(list_file_name, root):
    for file_name in list_file_name:
        path = root + "/" + file_name
        with pdfplumber.open(path) as pdf:
            p0 = pdf.pages[1]
            text = p0.extract_text(keep_blank_chars=True)
            output = open("output1.txt",'w')
            output.write(text)
            output = open('output1.txt', 'r')
            for num, line in enumerate(output.readlines()):
                myline = str(line).strip().split(' ')
                # if 'SERVICE' in myline and 'ADDRESS' in myline:
                #     month = myline[2] + '.'
                #     data['Month'].append(month)
                if 'DATE' in myline and 'MAILED' in myline and '$' in myline:
                    amount = myline[-1] 
                    data["Amount"].append(amount)
                if 'Therms' in myline and 'Used' in myline and 'THM' in myline:
                    usage = myline[-2]
                    data['Usage'].append(float(usage))
                if 'natural' in myline and 'gas' in myline and 'to' in myline and 'your' in myline and 'home':
                    meter = myline[-2]
                    meter= meter[0] + meter[-7:]
                    data['Meter'].append(meter)
    return data

def update_excel(workbook, new_values, month):
    specific_month = month
    sheet = workbook["Natural Gas"]
    
    for count, cell in enumerate(sheet[4]):
        if cell.value is not None:
            value = cell.value.split(" ")[0]
            if value == specific_month:
                index_amount = count - 1
                index_usage = count 
    
    for i, j in new_values.iterrows():
        for row in sheet.iter_rows(min_row=4):
            if row[4].value == j.iloc[1]:
                row[index_amount].value = j.iloc[0]
                row[index_usage].value = j.iloc[2]

    workbook.save('2024-25 Utility Billing.xlsx')

if __name__ == "__main__":
    #parse arguments
    parser = argparse.ArgumentParser("Parse info in pdf files")
    parser.add_argument("-f", help="folder name") 
    parser.add_argument("-e", help='existing excel file')
    parser.add_argument("-m", help="month") 
    args = parser.parse_args()

    files = os.listdir(args.f)
    dic_data = parse_pdf(files, args.f)

    new_values = pd.DataFrame(dic_data)
    workbook = openpyxl.load_workbook(args.e)
    update_excel(workbook, new_values, args.m)
    print('Successfuly updated the spreadsheet')



