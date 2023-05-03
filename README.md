# Comment Extractor for Word Documents

This Python script extracts comments from a Word document (.docx) and saves them in an Excel file (.xlsx). The script uses the python-docx and lxml libraries to parse the document file and extract the comments. The output file is saved in the same directory as the original document file.

## Requirements
- Python 3.x
- python-docx library
- lxml library
- pandas library
- xlsxwriter library
- google-colab library (for running in Google Colab environment)

## Usage
1. Install the required libraries: `pip install python-docx lxml pandas xlsxwriter google-colab`.
2. Clone the repository or download the script file (`comment_extractor.py`) and save it to your working directory.
3. In your Python script, import the `get_comments` function from `comment_extractor.py`.
4. Call the `get_comments` function with a bytes object of the Word document file as its argument.
5. The function returns a Pandas DataFrame object with the extracted comment data. You can manipulate or save this object as needed.

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments
This script was inspired by the `python-docx-demo` script from the python-docx library repository.
