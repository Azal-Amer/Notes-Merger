# Notes Merger
This Python script allows you to convert PowerPoint, Word, and Markdown files to PDF format, and then merge all the converted PDF files into a single file.
# Dependencies:

Python 3.x
PyPDF2
pypandoc
pandoc (command-line tool)
LibreOffice (with soffice command-line tool)

# Installing Dependencies:

- Install PyPDF2 using pip
- Install pypandoc using pip
- Install pandoc using your system's package manager
- Install LibreOffice using your system's package manager

# Usage:
- Clone the repository or download the script file
- Open a terminal or command prompt and navigate to the directory containing the script
- Run the script with the following command: python notesmerger.py
- By default, the script looks for files in the 'Notes to Merge' folder, converts them to PDF, and saves the converted files in the 'Converted Files' folder. It then merges all the PDF files (from both folders) into a single - file named 'merged.pdf'
- If you want to specify different folder paths or output file paths, modify the corresponding lines in the script

# Notes:

The script assumes that PowerPoint files have the .ppt or .pptx extension, Word files have the .docx extension, and Markdown files have the .md extension
The script recursively searches for files in the specified source folder and its subdirectories
After merging the PDF files, the converted folder is deleted to clean up the converted files