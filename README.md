# Worship List Maker

## Introduction

This program was created to help the ministry of worship director randomly assign members of the ministry to different days. This is so everyone gets a chance to participate and it saves time for the director to make the list.

NOTE: This program only works with Mac OS and Windows.

## Installation

Make sure you have Python(https://www.python.org/) and Microsoft Word(https://products.office.com/en-us/home) installed first. Then proceed with the instructions.

Download the project files to your computer, then run setup.py to install the necessary libraries.

To run setup.py, open up terminal and type in this command:

    $ cd
Then wherever your download folder is, drag it into the terminal. Hit enter.

Next, type this once the terminal knows what folder you want to be in:

    $ python setup.py
This will install the libraries that are needed to run the main program. Once setup.py is done, check out GoogleForm.txt to see how you should set up your google form to collect information.

Next, you'll need to have a formatted CSV file next to the main program which was downloaded from your Google Form.

To format the CSV file:

1. Inside Google Forms, create the spreadsheet.
2. Once spreadsheet is open, delete the first row of the data.
3. Next, delete the first column of the spreadsheet. Make sure the column you are deleting only contains timestamps. (i.e. numbers only)
4. Copy all of the data, then move it one column over to the left.
5. Download the file as a CSV file.
6. Open the CSV in any text editor.
7. Delete the first line which only has ",,,".
8. Save the file to the folder which has the main program in it.

Finally, type this in:

    $ python WorshipListMaker.py
The finished .docx file will be inside the same folder where the program is. Edit it in Microsoft Word for fine tuning.
## Aditional Notes

• To see the program in action, an example input has been included along with the program. Just rename the file to "input.csv".

• If in your data some people have a new cellphone number, a file called "newNumbers.csv" will appear in the same place as the newly created .docx file.

• To make lists for different months, change your system calendar to the month desired, then revert once done.
