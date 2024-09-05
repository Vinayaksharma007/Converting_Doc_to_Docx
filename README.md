**# Converting_Doc_to_Docx**
This script converts .doc files to .docx format using Microsoft Word via the comtypes library. It automates the process of converting older Microsoft Word documents to the modern .docx format.

**Requirements**
Python: Ensure you have Python installed (preferably Python 3.6 or higher).
Microsoft Word: The script requires Microsoft Word to be installed on your system.
comtypes: This library is used to interact with COM objects, like Microsoft Word. Install it using pip.

**Installation**
Install comtypes: You can install the comtypes library using pip. Open your command prompt or terminal and run:

**pip install comtypes**
Microsoft Word: Ensure that Microsoft Word is installed on your system, as the script uses it to perform the conversion.


**Script Configuration:**

Save the script to a file, e.g., convert_doc_to_docx.py.
Replace the doc_file_path in the script with the path to the .doc file you want to convert.

**Running the Script:**

Open your command prompt or terminal.

Navigate to the directory containing the script.

Run the script using Python:

python convert_doc_to_docx.py

Script Behavior:

The script will create a new .docx file in the same directory as the original .doc file.
It will print a success message with the path to the converted file or a failure message if the conversion did not succeed.

