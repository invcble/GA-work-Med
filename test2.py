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
    "Background Questions",
    "Admission",
    "Room",
    "Meals",
    "Nurses",
    "Tests and Treatments",
    "Visitors and Family",
    "Doctors",
    "Discharge",
    "Personal Issues",
    "Overall Assessment",
    "About You",
    "Comm w/ Nurses",
    "Response of Hosp Staff",
    "Comm w/ Doctors",
    "Hospital Environment",
    "Communication About Pain",
    "Comm About Medicines",
    "Discharge Information",
    "Care Transitions",
    "Comments",
    "Patient Name:"
    ]

start_index = filtered_text.find("Comm About Medicines")
end_index = filtered_text.find("Discharge Information")

# print(filtered_text[start_index : end_index])
print(filtered_text)
dict = {}

for each in section_headers:
    dict[each] = filtered_text.find(each)

clean_dict = {}
for each in dict:
    if dict[each] != -1:
        clean_dict[each] = dict[each]

# for i in dict:
#     print(dict[i], i)
for i in clean_dict:
    print(clean_dict[i], i)

##### ahahahahaahaha process dictionary to add next value in it, then itarate main list and reference
# list = list(clean_dict.items())
# print('-' * 40)
# # print(list[1][1])
# # print(len(list))
# for i in range(len(list)-1):
#     print(list[i][0])
#     print(f'{list[i][1]} to {list[i+1][1]}')

# workbook = openpyxl.Workbook()
# sheet = workbook.active

# for i in range(len(list)-1):
#     sheet.cell(row=5, column=i+5, value=list[i][0])
#     sheet.cell(row=6, column=i+5, value=f'{list[i][1]} to {list[i+1][1]}')
#     # print(list[i][0])
#     # print(f'{list[i][1]} to {list[i+1][1]}')

# workbook.save("Temp.xlsx")