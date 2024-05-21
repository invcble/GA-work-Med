import fitz  # PyMuPDF

def extract_text_with_bold(pdf_file):
    whole_text = ""
    doc = fitz.open(pdf_file)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"]
                        if "bold" in span["font"].lower():  # Check if the text is bold (flag 2 is bold in fitz)
                            whole_text += "(bold)" + text + "(bold)"
                            print("111111111111111111111111111111")
                        else:
                            whole_text += text
                    whole_text += "\n"
    
    return whole_text

pdf_file = "test.pdf"
text_with_bold = extract_text_with_bold(pdf_file)
print(text_with_bold)
