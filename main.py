import os
import re
import fitz
from PIL import Image
from reportlab.pdfgen import canvas
from datetime import datetime
from utils import progress_callback

def get_unique_number(pdf_name):
    unique_number_match = re.search(r'\d{9}', pdf_name)
    return unique_number_match.group() if unique_number_match else None

def convert_image_to_pdf(image_path, pdf_path):
    img = Image.open(image_path)
    pdf_path_with_extension = pdf_path if pdf_path.lower().endswith('.pdf') else pdf_path + '.pdf'
    c = canvas.Canvas(pdf_path_with_extension, pagesize=img.size)
    c.drawImage(image_path, 0, 0, width=img.size[0], height=img.size[1])
    c.save()
    return pdf_path_with_extension

def convert_tif_to_pdf(tif_path, pdf_path):
    image = Image.open(tif_path)
    pdf_path_with_extension = pdf_path if pdf_path.lower().endswith('.pdf') else pdf_path + '.pdf'
    image.save(pdf_path_with_extension, 'PDF', resolution=100.0)
    return pdf_path_with_extension

def extract_document_type(file_name):
    match = re.match(r'(\d{9}) - (.+?) - \d{1,2}_\d{1,2}_\d{4} - Stm\. Date - (\d{1,2}_\d{1,2}_\d{4}) - (\d{9})', file_name)
    if match:
        unique_number, doc_type, stm_date, _ = match.groups()
        return doc_type, unique_number, stm_date
    return None

def extract_date(file_name):
    match = re.search(r'Stm\. Date - (\d{1,2}_\d{1,2}_\d{4})', file_name)
    if match:
        return datetime.strptime(match.group(1), '%m_%d_%Y')
    return None

def is_skipped_file(file_name):
    skipped_extensions = ['.eml', '.htm', '.xlsx']
    return any(file_name.lower().endswith(ext) for ext in skipped_extensions)

def get_sort_key(file_name):
    order = [
        "Terms",
        "Bill Statement - CHARGE OFF",
        "Bill Statement",
        "Bill Statement - ACTIVITY",
        "Pay History",
        "Sales Memo",
        "Supplemental Borrower",
        "GOODBYE"
    ]
    for index, keyword in enumerate(order):
        if keyword in file_name:
            return index
    return len(order)  # If no keyword is found, place it at the end

def move_files_to_front(sorted_document_list, order):
    for keyword in reversed(order):
        for doc in sorted_document_list:
            if keyword in doc[3]:
                sorted_document_list.remove(doc)
                sorted_document_list.insert(0, doc)
    return sorted_document_list

def main(input_folder, output_folder, progress_queue):
    pdf_dict = {}

    for root, _, files in os.walk(input_folder):
        for file in files:
            if is_skipped_file(file):
                continue  # Skip files with specific extensions

            if file.lower().endswith(('.pdf', '.tif', '.tiff', '.jpg', '.jpeg')):
                unique_number = get_unique_number(file)
                if unique_number:
                    document_info = extract_document_type(file)
                    document_type = document_info[0] if document_info else None
                    file_path = os.path.join(root, file)

                    if file.lower().endswith(('.tif', '.tiff')):
                        pdf_path = convert_tif_to_pdf(file_path, file_path.replace('.tif', '_converted.pdf').replace('.tiff', '_converted.pdf'))
                        pdf_dict.setdefault(unique_number, []).append((document_type, pdf_path, extract_date(file), file))
                    elif file.lower().endswith(('.jpg', '.jpeg')):
                        pdf_path = convert_image_to_pdf(file_path, file_path.replace('.jpg', '_converted.pdf').replace('.jpeg', '_converted.pdf'))
                        pdf_dict.setdefault(unique_number, []).append((document_type, pdf_path, extract_date(file), file))
                    else:
                        pdf_dict.setdefault(unique_number, []).append((document_type, file_path, extract_date(file), file))

    total_files = sum(len(doc_list) for doc_list in pdf_dict.values())
    processed_files = 0

    for unique_number, document_list in pdf_dict.items():
        # Filter out documents without a valid date
        valid_documents = [doc for doc in document_list if doc[2] is not None]
        invalid_documents = [doc for doc in document_list if doc[2] is None]

        # Sort valid documents by custom order and date in ascending order
        valid_documents.sort(key=lambda x: (get_sort_key(x[3]), x[2] if x[2] else datetime.min))

        # Combine valid and invalid documents (invalid documents will be at the end)
        sorted_document_list = valid_documents + invalid_documents

        # Move files to the front based on the order list
        order = [
            "Terms",
            "Bill Statement - CHARGE OFF",
            "Bill Statement",
            "Bill Statement - ACTIVITY",
            "Pay History",
            "Sales Memo",
            "Supplemental Borrower",
            "GOODBYE"
        ]
        sorted_document_list = move_files_to_front(sorted_document_list, order)

        print(f"Order of files for unique number {unique_number}:")
        for doc in sorted_document_list:
            print(doc[3])

        merged_pdf = fitz.open()

        for _, document_path, _, _ in sorted_document_list:
            pdf_document = fitz.open(document_path)
            merged_pdf.insert_pdf(pdf_document)
            processed_files += 1
            process = processed_files / total_files
            progress_callback(progress_queue, process)

        output_filename = os.path.join(output_folder, f"{unique_number}.pdf")

        if merged_pdf.page_count > 0:
            merged_pdf.save(output_filename)
            merged_pdf.close()
            print(f"PDF {output_filename} merged successfully.")
        else:
            print(f"No pages found for {unique_number}.pdf. Skipping.")

    os.startfile(output_folder)

if __name__ == "__main__":
    input_folder = r''
    output_folder = r''

    # Test folders
    input_folder = r"P:\Users\Justin\Projects\output_test\rsg_doc_merger\01"
    output_folder = r"P:\Users\Justin\Projects\output_test\rsg_doc_merger\out"

    main(input_folder, output_folder)

    # Open File Explorer at the output_folder
    os.startfile(output_folder)

    print("PDFs merged successfully.")
