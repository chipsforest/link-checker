# link-checker
A Python script that checks the status of URLs in an Excel file and saves the results in a new Excel file.

Check Links

The check_links.py script is a Python script that checks the status of URLs in an Excel file and generates a report in a new Excel file.

Prerequisites

To use this script, you will need:

Python 3.5 or higher
Pandas
XlsxWriter
xlrd
requests
tqdm
Usage

To use the check_links.py script, follow these steps:

Install the required libraries using pip: pip install pandas xlsxwriter xlrd requests tqdm.
Save your Excel file containing URLs to check in the same directory as check_links.py.
Run the script in the command line: python check_links.py [filename.xlsx]. Replace [filename.xlsx] with the name of your Excel file.
Wait for the script to finish checking the URLs. This may take some time, depending on how many URLs are in your Excel file.
The script will generate a new Excel file called results.xlsx in the same directory as check_links.py. This file will contain a list of all URLs in the original Excel file along with their status codes. URLs with a status code of 404 will be highlighted in red.
Notes

The script assumes that the URLs are in the first sheet of the Excel file. If your URLs are in a different sheet, you will need to modify the script accordingly.
The script uses the tqdm library to display a progress bar during the URL checking process. If you prefer not to use this library, you can remove the tqdm import and the progress bar code from the script.
The script requires the XlsxWriter library to generate the output Excel file. If you prefer to use a different library or file format, you will need to modify the script accordingly.
If you encounter any issues or have suggestions for improvements, please open an issue on GitHub.
License

This script is released under the MIT License. See the LICENSE file for details.
