import re
import os
import sys
# import argparse
from docx import Document
from datetime import date
from docx2pdf import convert

def replace_string(filename):
    doc = Document(filename)
    for p in doc.paragraphs:
        if '${company}' in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if '${company}' in inline[i].text:
                    text = inline[i].text.replace('${company}', company)
                    inline[i].text = text
            print (p.text)
        
        if '${date}' in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if '${date}' in inline[i].text:
                    text = inline[i].text.replace('${date}', date)
                    inline[i].text = text
            print (p.text)

        if '${state}' in p.text:
            # print("FOUND STATE")
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if '${state}' in inline[i].text:
                    text = inline[i].text.replace('${state}', state)
                    inline[i].text = text
            print (p.text)

    docx_path = "./docx/CoverLetter_" + name + ".docx"
    doc.save(docx_path)

    pdf_path = "./PDF/CoverLetter_" + name + ".pdf"
    convert(docx_path, pdf_path)

try:
    arg = str(sys.argv)
    name = sys.argv[2]
    state = sys.argv[3]
    company = sys.argv[1]

    assert (len(sys.argv) == 4), "Expected: command line args: \"company\" name state"

except IndexError:
    raise SystemExit(f"Usage: {str(sys.argv)} \nExpected: command line args: \"company\" name state")

today = date.today()
# Textual month, day and year
date = today.strftime("%B %d, %Y")
print(company, state, date, "\n**\n")

replace_string("./coverletterdraft.docx")


