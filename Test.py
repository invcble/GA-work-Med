import pdfplumber
import re
import os
import sys


# path = "C:\\Users\\Administrator\\Downloads\\GA work Med"

if getattr(sys, 'frozen', False):
    current_path = os.path.dirname(sys.executable)
else:
    current_path = os.getcwd()

print("Current working path is", current_path)
pdf_files = []

for file in os.listdir(current_path):
    if file.endswith(".pdf"):
        pdf_files.append(file)

print(len(pdf_files), "PDFs found!")
print("Processing PDFs, Please wait!")


knownIDs = ["IZ0101", "IZ0101U", "HZ0101UE", "HZ0101U", "HZ0101"]
survey_count = {}

for pdf_file in pdf_files:
    
    wholeText = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:

            wholeText += page.extract_text()

    surveyArray = wholeText.split("Client Name:")
    surveyArray.pop(0)

    for survey in surveyArray:
        match = re.search("Survey Designator: (\\S+)", survey)
        # if match.group(1) not in knownIDs:
        #     print("New survey")
        survey_type = match.group(1)

        if survey_type in survey_count:
            survey_count[survey_type] += 1
        else:
            survey_count[survey_type] = 1

print(survey_count)
input("Press Enter to Exit....")