import sys
import os

import win32com.client
from pathlib import Path
from tqdm.auto import tqdm


def docToPDF(doc_filename, pdf_filename):
    """Export Microsoft Word document to PDF by pywin32, only Windows machine compatible
    
    Arguments:
        docx_filename {string} -- Input docx file name
        pdf_filename {string} -- Output pdf file name
    """
    # Make path absolute
    doc_path = Path(str(doc_filename)).resolve()
    pdf_path = Path(str(pdf_filename)).resolve()

    # File format for export from word (17 for PDF)
    export_format = 17

    # Load Microsoft Word application
    try:
        word = win32com.client.Dispatch('Word.Application')
    except:
        print('Microsoft Word not found on your computer.')
        sys.exit(0)

    # Open document
    try:
        doc = word.Documents.Open(str(doc_path))
    except:
        print('Error open document {filename}.'.format(filename=doc_filename))
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
        print('Export file error {filename}.'.format(filename=pdf_filename))
        sys.exit(0)

    # Close document and quit application
    doc.Close()
    word.Quit()

def batchDocToPDF(input_folder, output_folder):
    """Batch export Mictosoft Word cocument to PDF
    
    Arguments:
        input_folder {string} -- Input folder path
        output_folder {string} -- Output folder path
    """
    # Make folder path absolute
    in_folder_path = Path(str(input_folder)).resolve()
    out_folder_path = Path(str(output_folder)).resolve()

    # Export every .doc or .docx document to PDF
    for doc_path in tqdm(sorted(Path(in_folder_path).glob('[0-9a-zA-Zа-яА-Я+-]*.doc*'))):
        # PDF file path
        pdf_path = Path(out_folder_path) / (str(doc_path.stem) + '.pdf')

        # Check dir existing and export document to PDF
        if Path.is_dir(out_folder_path):
            docToPDF(doc_path, pdf_path)
        else:
            print(' Folder doen`t exist, creating new one.')
            os.mkdir(Path(output_folder).resolve())
            docToPDF(doc_path, pdf_path)