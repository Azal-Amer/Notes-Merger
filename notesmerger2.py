import os
import subprocess
import pypandoc
from PyPDF2 import PdfMerger
import shutil
from weasyprint import HTML, CSS
from weasyprint import HTML, CSS

def convert_md_to_pdf(md_folder, converted_folder):
    md_folder = os.path.abspath(md_folder)
    converted_folder = os.path.abspath(converted_folder)

    # Create the converted folder if it doesn't exist
    os.makedirs(converted_folder, exist_ok=True)

    try:
        # Convert Markdown files to PDF
        md_files = [os.path.join(root, file) for root, _, files in os.walk(md_folder) for file in files
                    if file.endswith('.md') and not file.startswith('~$')]

        for md_file in md_files:
            pdf_file = os.path.splitext(os.path.basename(md_file))[0] + '.pdf'
            pdf_path = os.path.join(converted_folder, pdf_file)

            with open(md_file, 'r') as f:
                html = pypandoc.convert_text(f.read(), 'html', format='md')

            css = CSS(string='''
                @page { size: A4; margin: 1cm; }
                body { font-family: Arial, sans-serif; font-size: 12pt; }
            ''')

            HTML(string=html).write_pdf(pdf_path, stylesheets=[css])

        print(f"Successfully converted {len(md_files)} Markdown files to PDFs in '{converted_folder}'.")
    except Exception as e:
        print(f"An error occurred during Markdown conversion: {str(e)}")
def convert_files_to_pdf(source_folder, converted_folder):
    source_folder = os.path.abspath(source_folder)
    converted_folder = os.path.abspath(converted_folder)

    # Create the converted folder if it doesn't exist
    os.makedirs(converted_folder, exist_ok=True)

    try:
        # Convert PowerPoint files to PDF
        # ppt_files = [os.path.join(source_folder, file) for file in os.listdir(source_folder)
        #              if file.endswith(('.ppt', '.pptx')) and not file.startswith('~$')]
        ppt_files = [os.path.join(root, file) for root, _, files in os.walk(source_folder) for file in files
             if file.endswith(('.ppt', '.pptx')) and not file.startswith('~$')]

        for ppt_file in ppt_files:
            pdf_file = os.path.splitext(os.path.basename(ppt_file))[0] + '.pdf'
            pdf_path = os.path.join(converted_folder, pdf_file)

            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', converted_folder, ppt_file])

        # Convert Word files to PDF
        # docx_files = [os.path.join(source_folder, file) for file in os.listdir(source_folder)
        #               if file.endswith('.docx') and not file.startswith('~$')]
        docx_files = [os.path.join(root, file) for root, _, files in os.walk(source_folder) for file in files
              if file.endswith('.docx') and not file.startswith('~$')]

        for docx_file in docx_files:
            pdf_file = os.path.splitext(os.path.basename(docx_file))[0] + '.pdf'
            pdf_path = os.path.join(converted_folder, pdf_file)

            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', converted_folder, docx_file])

        # Convert Markdown files to PDF
        # md_files = [os.path.join(source_folder, file) for file in os.listdir(source_folder)
        #             if file.endswith('.md') and not file.startswith('~$')]
        md_files = [os.path.join(root, file) for root, _, files in os.walk(source_folder) for file in files
            if file.endswith('.md') and not file.startswith('~$')]

        for md_file in md_files:
            pdf_file = os.path.splitext(os.path.basename(md_file))[0] + '.pdf'
            pdf_path = os.path.join(converted_folder, pdf_file)
            

            output = pypandoc.convert_file(md_file, 'pdf', outputfile=pdf_path)

        print(f"Successfully converted files to PDFs in '{converted_folder}'.")
    except Exception as e:
        print(f"An error occurred during file conversion: {str(e)}")

def merge_pdfs(source_folder, converted_folder, output_path):
    # Create a PdfMerger object
    merger = PdfMerger()

    # Get a list of all PDF files in the source folder and the converted folder
    pdf_files = [os.path.join(root, file) for root, _, files in os.walk(source_folder) for file in files if file.endswith('.pdf')]
    pdf_files.extend([os.path.join(root, file) for root, _, files in os.walk(converted_folder) for file in files if file.endswith('.pdf')])

    # Sort the PDF files alphabetically
    pdf_files.sort()

    try:
        # Iterate over each PDF file and append it to the merger
        for pdf_file in pdf_files:
            file_path = os.path.join(source_folder, pdf_file) if os.path.exists(os.path.join(source_folder, pdf_file)) else os.path.join(converted_folder, pdf_file)
            merger.append(file_path)

        # Write the merged PDF to the output file
        merger.write(output_path)
        merger.close()

        print(f"Successfully merged {len(pdf_files)} PDFs into '{output_path}'.")
    except Exception as e:
        print(f"An error occurred during PDF merging: {str(e)}")
    finally:
        # Delete the converted folder
        shutil.rmtree(converted_folder)
        print(f"Deleted '{converted_folder}'.")

# Specify the folder path containing the source files
source_folder = input('Enter the folder path containing the source files: ')

# Specify the folder path for saving converted files
converted_folder = source_folder+  '/converted_files'

# Specify the output file path for the merged PDF
output_path = source_folder+'merged.pdf'

# Convert source files to PDFs
convert_files_to_pdf(source_folder, converted_folder)

# Merge all PDFs in the source folder and the converted folder
merge_pdfs(source_folder, converted_folder, output_path)