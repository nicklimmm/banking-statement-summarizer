# Welcome to banking-statement-summarizer repo

:information_source: *Note: This summarizer only works with Frank OCBC Account as of now*

:warning: *Please use this summarizer with caution! The script is intended to be used **personally** ([Why is this safe?](#why-is-this-safe?)). I will not be held responsible for the leakage of private information into the public.*

## Why is this repo created?
I found a hard time tracking my cashflow and since I receive my banking e-statements every month, I thought to make a summary and visualisation from all statements that I got. But, the e-statements are in PDF form and as I receive more e-statements in the future, it would be tedious to extract the relevant information into an Excel sheet and create visualizations one-by-one manually.

So, I came up with a brilliant solution (I think so :/), which is to make use of Python to automate the process. And I am looking forward for others to contribute on this project to cover other banks, so that everyone can make use of this automation for their finances.

## For whom is this repo intended?
Perfect for anyone who wants to track their cashflow by creating a summary and visualisation of it.

## Installation
Make sure you have Python 3 installed in your machine. If you haven't, you can download from [Python Website](https://www.python.org/downloads/) or you could find tutorials in Youtube on how to install Python 3.
- Using zip file
1. Download the zip file from this repo and extract the folder to a location that is familiar to you (e.g. Desktop). Open that folder.
2. For **Windows** users, type *cmd* in the location bar of Windows Explorer and press Enter. For **Linux** users, right-click on an empty space and click 'Open in Terminal'. For **Mac** users, detailed steps coming soon...
3. Type or paste this code in the cmd or terminal
   ```
   pip install -r requirements.txt
   ``` 
   or 
   ```
   pip3 install -r requirements.txt
   ```
   , this command will install the required Python packages to run the script.

- Using git (Make sure that git is installed, if haven't, install it from [git website](https://git-scm.com/).)
1. For **Windows** users, type *cmd* in the location bar of Windows Explorer and press Enter. For **Linux** users, right-click on an empty space and click 'Open in Terminal'. For **Mac** users, detailed steps coming soon...
2. In cmd or terminal, change your current working directory to somewhere that you prefer (e.g. Desktop or Downloads) using the ```cd``` command.
3. Type or paste this code in the cmd or terminal
   ```
   git clone https://github.com/nicklimmm/banking-statement-summarizer
   ```
4. Type or paste this code in the cmd or terminal
   ```
   pip install -r requirements.txt
   ``` 
   or 
   ```
   pip3 install -r requirements.txt
   ```
   , this command will install the required Python packages to run the script.

## How can I use this?
1. Open the script folder, then move all your PDF e-Statements into it (Make sure to not rename the original e-Statements).
2. On that folder, for **Windows** users, type *cmd* in the location bar of Windows Explorer and press Enter. For **Linux** users, right-click on an empty space and click 'Open in Terminal'. For **Mac** users, detailed steps coming soon...
3. Type or paste this code in the cmd or terminal
   ```
   python3 create_summary.py
   ```
   or
   ```
   python create_summary.py
   ```
   or
   ```
   py create_summary.py
   ```
   , this command will run the summarizer script.
4. In order to have access on your PDF e-statement, the script will ask for the password ([Why is this safe?](#why-is-this-safe?)).
5. Wait for a few seconds, then an Excel file will be created in the same folder that contains the summary and the visualization of your cashflow.

## Why is this safe?
- When the script is asking for your password, the password will be hidden when you type the password.
- The script does not save your passwords in any way.
- The Excel file created only contains the cashflow.
- After unlocking your PDF e-Statements, it will be locked again.
