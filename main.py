import os
import re
import fitz
import openpyxl
from PIL import Image
from reportlab.pdfgen import canvas
from datetime import datetime
from utils import progress_callback

ORDER = [
    "Terms",
    "Supplemental Borrower",
    "Bill Statement - CHARGE OFF",
    "Bill Statement",
    "Pay History",
    "Sales Memo",
    "GOODBYE",
    "OwnershipChain",
    "DataString",
]

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
    match = re.match(
        r'(\d{9}) - (.+?) - \d{1,2}_\d{1,2}_\d{4} - Stm\. Date - (\d{1,2}_\d{1,2}_\d{4}) - (\d{9})',
        file_name,
    )
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
    """
    Return integer index according to ORDER (case-insensitive).
    Files that don't match any keyword get index == len(ORDER).
    """
    lower_name = (file_name or "").lower()
    for index, keyword in enumerate(ORDER):
        if keyword.lower() in lower_name:
            return index
    return len(ORDER)

def move_files_to_front(sorted_document_list, order=ORDER):
    """
    Return a new list where documents matching 'order' keywords are
    grouped at the front in the order defined by 'order'.
    Each group's internal order is preserved; unmatched items follow.
    """
    groups = {i: [] for i in range(len(order))}
    remainder = []
    for doc in sorted_document_list:
        idx = get_sort_key(doc[3])
        if idx < len(order):
            groups[idx].append(doc)
        else:
            remainder.append(doc)
    result = []
    for i in range(len(order)):
        result.extend(groups[i])
    result.extend(remainder)
    return result

def check_pdf_integrity(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            _ = doc.page_count
        return True
    except Exception:
        return False

def write_merge_results_to_excel(results, excel_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["rsg number", "merged"])
    for unique_number, merged in results:
        ws.append([unique_number, "x" if merged else ""])
    wb.save(excel_path)

def main(input_folder, output_folder, progress_queue=None):
    pdf_dict = {}
    merge_results = []

    for root, _, files in os.walk(input_folder):
        for file in files:
            if is_skipped_file(file):
                continue

            if file.lower().endswith(('.pdf', '.tif', '.tiff', '.jpg', '.jpeg')):
                unique_number = get_unique_number(file)
                if not unique_number:
                    continue

                document_info = extract_document_type(file)
                document_type = document_info[0] if document_info else None
                file_path = os.path.join(root, file)

                base_no_ext = os.path.splitext(file_path)[0]
                if file.lower().endswith(('.tif', '.tiff')):
                    pdf_path = convert_tif_to_pdf(file_path, base_no_ext + '_converted.pdf')
                    pdf_dict.setdefault(unique_number, []).append((document_type, pdf_path, extract_date(file), file))
                elif file.lower().endswith(('.jpg', '.jpeg')):
                    pdf_path = convert_image_to_pdf(file_path, base_no_ext + '_converted.pdf')
                    pdf_dict.setdefault(unique_number, []).append((document_type, pdf_path, extract_date(file), file))
                else:
                    pdf_dict.setdefault(unique_number, []).append((document_type, file_path, extract_date(file), file))

    total_files = sum(len(doc_list) for doc_list in pdf_dict.values()) or 1
    processed_files = 0

    for unique_number, document_list in pdf_dict.items():
        valid_documents = [doc for doc in document_list if doc[2] is not None]
        invalid_documents = [doc for doc in document_list if doc[2] is None]

        valid_documents.sort(key=lambda x: (get_sort_key(x[3]), x[2] if x[2] else datetime.min))
        sorted_document_list = valid_documents + invalid_documents
        sorted_document_list = move_files_to_front(sorted_document_list)

        print(f"Order of files for unique number {unique_number}:")
        for doc in sorted_document_list:
            print(doc[3])

        all_files_ok = True
        for _, document_path, _, _ in sorted_document_list:
            if not check_pdf_integrity(document_path):
                print(f"Corrupted file detected: {document_path}. Skipping merge for {unique_number}.")
                all_files_ok = False
                break

        merged_pdf = fitz.open()
        output_filename = os.path.join(output_folder, f"{unique_number}.pdf")

        if all_files_ok:
            for _, document_path, _, _ in sorted_document_list:
                pdf_document = fitz.open(document_path)
                merged_pdf.insert_pdf(pdf_document)
                pdf_document.close()
                processed_files += 1
                progress = processed_files / total_files
                progress_callback(progress_queue, progress)

            if merged_pdf.page_count > 0:
                merged_pdf.save(output_filename)
                merged_pdf.close()
                print(f"PDF {output_filename} merged successfully.")
                merge_results.append((unique_number, True))
            else:
                merged_pdf.close()
                print(f"No pages found for {unique_number}.pdf. Skipping.")
                merge_results.append((unique_number, False))
        else:
            merge_results.append((unique_number, False))

    excel_path = os.path.join(output_folder, "merge_results.xlsx")
    write_merge_results_to_excel(merge_results, excel_path)

    try:
        os.startfile(output_folder)
    except Exception:
        pass

if __name__ == "__main__":
    input_folder = r"P:\Users\Steven Cox\Projects\test folders\Rsg Doc merger\In"
    output_folder = r"P:\Users\Steven Cox\Projects\test folders\Rsg Doc merger\Out"

    main(input_folder, output_folder)

    print("PDFs merged successfully.")