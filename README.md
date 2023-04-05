# Bitbucket Server Commit Extractor

This script is used to extract information about all the commits in all the branches of all the repositories in a Bitbucket Server project and write the information to an Excel file. The script can be run from the command line and requires the following inputs:

- The Bitbucket Server API URL and project key
- Bitbucket Server credentials for authentication
- The path of the output Excel file

## Requirements

- Python 3.x
- requests
- openpyxl

## Usage

To use this script, follow these steps:

1. Install the required packages by running the following command:

```bash
pip3 install requests openpyxl
```

2. Download the script and save it to a local directory.

3. Open the script in a text editor and modify the following variables at the beginning of the script:

- `API_URL`: Set the Bitbucket Server API URL and project key.
- `USERNAME` and `PASSWORD`: Set the Bitbucket Server credentials for authentication.
- `EXCEL_FILE`: Set the path of the output Excel file.

4. Save the modified script.

5. Open a terminal or command prompt and navigate to the directory where the script is saved.

6. Run the script by entering the following command:

The script will start running and will output a message indicating the progress of the extraction process.

When the script completes, open the output Excel file to view the extracted commit information.

## Output

The output of the script is an Excel file containing a table with the following columns for each commit:

- Repository: the name of the repository
- Branch: the name of the branch
- Commit: the ID of the commit
- Author: the name of the author who made the commit
- Date: the date and time the commit was made
- Message: the commit message

If the commit message contains illegal characters that cannot be written to the Excel file, the script writes an empty string to the corresponding cell.

