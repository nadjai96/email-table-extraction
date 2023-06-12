# Email Table Extraction

This Python project extracts tables from your Outlook emails and saves them into an Excel file.

## Project Description

The script scans through your emails, filters emails based on a specified subject and a certain number of past days from the current date. For each matching email, it extracts the first HTML table it encounters and writes it to an Excel file. Each sheet in the Excel file corresponds to an email with the timestamp as the sheet name.

The script has the ability to filter identical neighboring columns, ensuring no redundant information is captured.

## Prerequisites

Ensure you have the following installed on your machine:

- Python 3
- `pandas` library
- `pytz` library
- `pywin32` library

You can install the libraries using pip:

```bash
pip install pandas pytz pywin32

Note: This script requires access to your Outlook client and will read from your Inbox. Please ensure that you have necessary permissions and access rights.

Setup and Usage

Clone the repository:

git clone https://github.com/<your_username>/email-table-extraction.git

Navigate to the project directory:

cd email-table-extraction
Open the script in your preferred text editor or IDE.

Replace Your_Output_Folder in the script with the path to the folder where you want to save your files. Make sure that the path is absolute.

Replace Your_Subject_String with the subject string that you want to check for in the subject of the emails.

Replace num_days with the number of days from the current date for which you want to check the emails.

Save the script and run it using a Python interpreter.

License

Distributed under the MIT License.