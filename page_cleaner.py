#pip install PyMuPDF <<<< might need to install the pkg

import os
import fitz  # PyMuPDF

def delete_first_two_pages(input_pdf, output_pdf):
    doc = fitz.open(input_pdf)
    
    # Delete the first two pages
    if len(doc) > 2:
        doc.delete_pages(0, 1)
    elif len(doc) == 2:
        doc.delete_page(1)
        doc.delete_page(0)
    elif len(doc) == 1:
        doc.delete_page(0)
    

    doc.save(output_pdf)
    doc.close()


def remove_pages_with_text(input_pdf, output_pdf, search_text):
    doc = fitz.open(input_pdf)
    pages_to_delete = []


    for page_num in range(len(doc)):
        text = doc[page_num].get_text("text").strip()
        if text.startswith(search_text):
            pages_to_delete.append(page_num)
    
    # Delete pages in reverse order to avoid shifting indices
    for page_num in reversed(pages_to_delete):
        doc.delete_page(page_num)


    doc.save(output_pdf)
    doc.close()


dir_input = '/Users/xuan/Downloads/Pilot_data/Original/'
dir_middle = '/Users/xuan/Downloads/Pilot_data/test/'
dir_output ='/Users/xuan/Downloads/Pilot_data/cleaned/'
file_list = os.listdir(dir_input)


for j in file_list:
    if j != '.DS_Store':
        remove_pages_with_text(dir_input+j, dir_middle+j, 'Graded')

for j in file_list:
    if j != '.DS_Store':
        remove_pages_with_text(dir_middle+j, dir_output+j, 'Question assigned to the following page: 1') 
