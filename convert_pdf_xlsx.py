"""
Original author: rkf33
Updated by chenxvan for Mac

python3 -m venv myenv
source myenv/bin/activate

possible pkg needed:
pip install pytesseract pdf2image tqdm openpyxl pandas nltk
"""
import os
import pandas as pd
import nltk
from pdf2image import convert_from_path
import pytesseract
from tqdm import tqdm

# Ensure nltk punkt tokenizer is available
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')


def pdf_to_text(pdf_path):
    try:
        images = convert_from_path(pdf_path)

        full_text = ""
        for image in images:
            text = pytesseract.image_to_string(image)
            text = text.replace('|', 'I') 
            full_text += text + "\n"

        return full_text

    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return ""


def txt_to_xlsx(txt_file_path, output_folder):
    try:
        with open(txt_file_path, "r", encoding="utf-8") as file:
            text = file.read()

        raw_text = []
        raw_text_context = []

        paragraphs = text.split('\n\n')
        for paragraph in paragraphs:
            clean_paragraph = paragraph.strip()
            if clean_paragraph:
                for sentence in nltk.sent_tokenize(paragraph):
                    
                    raw_text.append(sentence)

                    raw_text_context.append(clean_paragraph)

        output_xlsx = os.path.join(output_folder, os.path.splitext(os.path.basename(txt_file_path))[0] + '.xlsx')
        with pd.ExcelWriter(output_xlsx) as writer:
            pd.DataFrame({'Sentences': raw_text}).to_excel(
                writer,
                sheet_name='Sentence Data')

        print(f" Saved XLSX: {output_xlsx}")

    except Exception as e:
        print(f"Error converting {txt_file_path} to XLSX: {e}")

def process_folder(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]

    for pdf_file in tqdm(pdf_files, desc="Processing PDFs"):
        pdf_path = os.path.join(input_folder, pdf_file)
        output_txt_path = os.path.join(output_folder, os.path.splitext(pdf_file)[0] + '.txt')

        text = pdf_to_text(pdf_path)
        if text.strip():
            with open(output_txt_path, 'w', encoding='utf-8') as f:
                f.write(text)

            # After saving TXT, convert it to XLSX
            txt_to_xlsx(output_txt_path, output_folder)

    print(f"Finished processing {len(pdf_files)} PDFs. Output saved in: {output_folder}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert scanned PDFs to XLSX.")
    parser.add_argument("input_folder", help="Folder containing scanned PDF files.")
    parser.add_argument("output_folder", help="Folder to save extracted TXT and XLSX files.")

    args = parser.parse_args()

    process_folder(args.input_folder, args.output_folder)
