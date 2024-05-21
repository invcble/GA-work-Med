import pdfplumber
import re
import os
import sys
import openpyxl



wholeText = ""

with pdfplumber.open("test.pdf") as pdf:
    for page in pdf.pages:
        wholeText += page.extract_text()


filtered_text = wholeText.replace("© 2023 Press Ganey Associates LLC † Custom Question ^ Focus Question", "")

# print(filtered_text[(filtered_text.find("Care Transitions")):])

# print(filtered_text.find("Care Transitionss"))

section_headers = [
    ["Background Questions", 0, 0],
    ["Admission", 0, 0],
    ["Room", 0, 0],
    ["Meals", 0, 0],
    ["Nurses", 0, 0],
    ["Tests and Treatments", 0, 0],
    ["Visitors and Family", 0, 0],
    ["Doctors", 0, 0],
    ["Discharge", 0, 0],
    ["Personal Issues", 0, 0],
    ["Overall Assessment", 0, 0],
    ["About You", 0, 0],
    ["Comm w/ Nurses", 0, 0],
    ["Response of Hosp Staff", 0, 0],
    ["Comm w/ Doctors", 0, 0],
    ["Hospital Environment", 0, 0],
    ["Communication About Pain", 0, 0],
    ["Comm About Medicines", 0, 0],
    ["Discharge Information", 0, 0],
    ["Care Transitions", 0, 0],
    ["Comments", 0, 0],
    ["Patient Name:", 0, 0]
    ]

start_index = filtered_text.find("Comm About Medicines")
end_index = filtered_text.find("Discharge Information")

# print(filtered_text[start_index : end_index])
# print(filtered_text)
dict = {}

for each in section_headers:
    if filtered_text.find(each) != -1:
        dict[each] = filtered_text.find(each)

for i in dict:
    print(dict[i], i)

##### ahahahahaahaha process dictionary to add next value in it, then itarate main list and reference
list = list(dict.items())
print('-' * 40)
# print(list[1][1])
# # print(len(list))
# print(list)
for i in range(len(list)-1):
    print(list[i][0])
    print(f'{list[i][1]} to {list[i+1][1]}')

workbook = openpyxl.Workbook()
sheet = workbook.active

for i in range(len(list)-1):
    sheet.cell(row=5, column=i+5, value=list[i][0])
    sheet.cell(row=6, column=i+5, value=f'{list[i][1]} to {list[i+1][1]}')
    # print(list[i][0])
    # print(f'{list[i][1]} to {list[i+1][1]}')

workbook.save("Temp.xlsx")