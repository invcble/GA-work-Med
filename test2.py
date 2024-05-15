import pdfplumber
import re
import os
import sys



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
    "Care Transitions"
    ]

start_index = filtered_text.find("Comm About Medicines")
end_index = filtered_text.find("Discharge Information")

print(filtered_text[start_index : end_index])