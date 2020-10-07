# excel-python
A repo that explores the openpyxl Python library

### Installation
1. Clone this repository
2. Create a new Python3 virtual environment using `python3 -m venv <SOME_ENV_NAME>`.
3. Activate the environment using the command `source <YOUR_ENV_NAME>/bin/activate`
4. Once the environment is activated, run `pip install -r requirements.txt`
5. After pip is finished installing you can begin to use any of the scripts in the repository.

### Using XL2Text.py
1. Ensure the Python virtual environment is activated
2. Run the script with `python XL2Text.py`
3. The program will prompt you to enter the full path to an excel file. You
can either add the excel file to the root directory of the repository and just enter the file name, or give a full path to a file elsewhere in the file system.
4. Optionally specify a column width 
5. The sheets in the excel file will be written to text files in the specified directory. If no directory is specified, the text files will be written to the same directory as XL2Text.py

