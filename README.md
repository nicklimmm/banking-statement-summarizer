# Welcome to banking-statement-summarizer repo

:information_source: *Note: This summarizer only works with Frank OCBC Account as of now*

:warning: *Please use this summarizer with caution! The script is intended to be used **personally**. I will not be held responsible for the leakage any private information into the public.*

## Why is this repo created?
I found a hard time tracking how much money do I get or spend in a certain amount of time and since I receive my banking e-statements every month, I thought to make a summary and visualisation from all statements that I got. 

But, the e-statements are in PDF form and as I receive more e-statements in the future, it would be tedious to extract the relevant information into an Excel sheet and create visualizations one-by-one manually. So, I came up with a brilliant solution (I think so :/), which is to make use of Python to automate the process.

## For whom is this repo intended?
Perfect for anyone who wants to track their cashflow by creating a summary and visualisation of it, and had some experience in using Python.

## Installation
Make sure you have Python 3 installed in your machine. If you haven't, you can download from [Python's Official Website](https://www.python.org/downloads/) or you could find tutorials in Youtube on how to install Python 3.
- **Windows**
1. Download the zip file from this repo and extract the folder to a location that is familiar to you (e.g. Desktop).
2. Open that folder, type *cmd* in the location bar of Windows Explorer and press Enter. The command prompt will be opened.
3. Type ```pip install -r requirements.txt```, this command will install the required Python packages to run the script.
- **MacOs and Linux**  *(Detailed steps coming soon... You can install according to your OS, if stuck, you could search for help in the internet)*

## How can I use this?
1. Open the script folder, then move all your PDF e-Statements into it (Make sure to not rename the original e-Statements).
2. On that folder, type *cmd* in the location bar of Windows Explorer and press Enter. The command prompt will be opened.
3. Type ```python create_summary.py```, this command will run the summarizer script.
4. In order to have access on your PDF e-statement, the script will ask for the password ([Why is this safe?](#why-is-this-safe?)).
5. Wait for a few seconds, then an Excel file will be created in the same folder that contains the summary and the visualization of your cashflow.

## Why is this safe?
- When the script is asking for your password, the password will be hidden when you type the password.
- The script does not save your passwords in any way.
- The Excel file created only contains the cashflow.
- After unlocking your PDF e-Statements, it will be locked again.