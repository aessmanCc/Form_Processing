Excel VBA Automation Project

This project consists of VBA macros embedded in Excel files (.xlsm) that automate various tasks related to data entry, saving, and processing in a business environment. The macros are designed to enhance productivity and streamline workflows for users interacting with Excel workbooks.

Functionality Overview

1. SaveWork Macro:
- When the save button is clicked, this macro prompts the user whether they want to process another task.
- If the user chooses to continue (yes), the current workbook is saved to a specified folder based on store information, and another master Excel file is opened for further operations.
- If the user chooses not to continue (no), the current workbook is simply saved and closed.

2. UserForm1:
- The UserForm1 provides a user-friendly interface for data entry.
- Users can enter information about products, including sales, revenue, costs, item description, ticket number, agreement date, etc.
- Upon submission, the entered data is transferred to specific cells in the Excel worksheet.

Usage

1. Save Button Click:
- Clicking the save button triggers the SaveWork macro, initiating the saving process and providing options for further tasks.

2. Data Entry with UserForm1:
- Clicking the data entry button opens UserForm1, where users can input product information conveniently.
- The entered data is then transferred to the Excel worksheet for processing and analysis.

File Structure

- Master Excel Files: The project includes two master Excel files, one of which is opened by the SaveWork macro.
- UserForm1.frx: This file contains the graphical layout and properties of the user form for data entry.

Prerequisites 

- Microsoft Excel installed
- Macros enabled in Excel settings
- Proper configuration of file paths and destination folders

Important Notes 

- Ensure that the file paths specified in the macros are accurate and accessible.
- Customize the macros and user form according to your specific business requirements.
- Test the functionality thoroughly to ensure proper operation and error handling.