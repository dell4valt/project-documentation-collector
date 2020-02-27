import sys
import os

import win32com.client
from pathlib import Path
from tqdm.auto import tqdm
import re
from PyPDF2  import PdfFileWriter, PdfFileReader, PdfFileMerger

def doc_to_pdf(doc_filename, pdf_filename):
    """Export Microsoft Word document to PDF by pywin32, only Windows machine compatible.
    
    Arguments:
        docx_filename {str} -- Input docx file name
        pdf_filename {str} -- Output pdf file name
    """
    # Make path absolute
    doc_path = Path(str(doc_filename)).resolve()
    pdf_path = Path(str(pdf_filename)).resolve()

    # File format for export from word (17 for PDF)
    export_format = 17

    # Load Microsoft Word application
    try:
        word = win32com.client.Dispatch("Word.Application")
    except:
        print("Microsoft Word not found on your computer.")
        sys.exit(0)

    # Open document
    try:
        doc = word.Documents.Open(str(doc_path))
    except:
        print("Error open document {filename}.".format(filename=doc_filename))
        sys.exit(0)

    # Export document to PDF with bookmarks and good quiality
    try:
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_path),
            ExportFormat=export_format,
            OptimizeFor=0,
            CreateBookmarks=1,
            DocStructureTags=True
        )
    except:
        print("Export file error {filename}.".format(filename=pdf_filename))
        sys.exit(0)

    # Close document and quit application
    doc.Close()
    word.Quit()


def batch_doc_to_pdf(input_folder, output_folder):
    """Batch export Mictosoft Word document to PDF.
    
    Arguments:
        input_folder {str} -- Input folder path
        output_folder {str} -- Output folder path
    """
    # Make folder path absolute
    in_folder_path = Path(str(input_folder)).resolve()
    out_folder_path = Path(str(output_folder)).resolve()

    # Export every .doc or .docx document to PDF
    for doc_path in tqdm(sorted(Path(in_folder_path).glob("[0-9a-zA-Zа-яА-Я+-]*.doc*"))):
        # PDF file path
        pdf_path = Path(out_folder_path) / (str(doc_path.stem) + ".pdf")

        # Check dir existing and export document to PDF
        if Path.is_dir(out_folder_path):
            doc_to_pdf(doc_path, pdf_path)
        else:
            print(" Folder doen`t exist, creating new one.")
            os.makedirs(Path(output_folder).resolve())
            doc_to_pdf(doc_path, pdf_path)


def collect(project_folder, out_filename):
    """Collect PDF files to main document.
    Function searches for PDF files whose names satisfy the conditions,
    and collect it to the main PDF document.
    The title of the Cover Page must contain a “титул” or “обложка”.
    The name of the Information and Certification Sheet should contain "УЛ", "ИУЛ", "Информационно-удостоверяющий лист".
    The name of the Change Registration Table should include a “таблица регистрации изменений".
    The title of the Main document of the sheet should contain “ПЗ” or “Пояснительная записка”.
    
    Arguments:
        project_folder {str} -- Path to directory with PDF files
    """
    project_folder = Path(str(project_folder)).resolve()

    # List of PDF files paths
    files_list = sorted(Path(project_folder).glob("*.pdf"))
    files_list_str = list(str(s) for s in files_list)

    # Regular expressions for determine files type
    info_cert_page_re = re.compile(r"УЛ|ИУЛ|Информационно-удостоверяющий лист|информационно-удостоверяющий лист")
    title_page_re = re.compile(r"титул|обложка", re.IGNORECASE)
    changes_page_re = re.compile(r"таблица регистрации изменений", re.IGNORECASE)
    main_doc_re = re.compile(r"ПЗ|Пояснительная записка|пояснительная записка")

    # Get files paths
    info_cert_path = list(filter(info_cert_page_re.search, files_list_str))
    title_path = list(filter(title_page_re.search, files_list_str))
    changes_path = list(filter(changes_page_re.search, files_list_str))
    main_doc_path = list(filter(main_doc_re.search, files_list_str))

    merger = PdfFileMerger()
    main_doc_pdf = PdfFileReader(open(main_doc_path[0], "rb"))

    # Append title pages if exist to main document, or just append main document
    if title_path:
        title_page_pdf = PdfFileReader(open(title_path[0], "rb"))
        title_page_num = title_page_pdf.getNumPages()

        main_doc_page_num = main_doc_pdf.getNumPages()
        merger.append(fileobj=title_page_pdf)
        print("Title page appended to document.")
        merger.merge(position=2, fileobj=main_doc_pdf, pages=(title_page_num, main_doc_page_num))
        print("Main doc appended to document.")
    else:
        merger.append(fileobj=main_doc_pdf)
        print("Main doc appended to document.")

    
    # Append Table of changes page to main document
    if changes_path:
        changes_page_pdf = PdfFileReader(open(changes_path[0], "rb"))
        merger.append(fileobj=changes_page_pdf)
        print("Change Registration Table page appended to document.")

        # Adding bookmark to inserted Table of changes page
        merger.addBookmark("Таблица регистрации изменений", main_doc_pdf.getNumPages())

    output = open(out_filename, "wb")
    merger.write(output)
    merger.close()
