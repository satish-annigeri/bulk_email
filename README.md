# `bulk_email`

`bulk_email` sends individual emails using Gmail SMTP to a list of recipients whose name, email address and other parameters are read from a Microsoft Excel file.

It uses Mako template library to implement processing logic and looping. Data of each recipient is stored as a tuple with the first item is the email address and the second is the name of the recipient. The rest can be any other data required by the template.

## Required

Following packages must be installed, preferably in a virtual environment.

1. `python-dotenv` to manage environment variables
2.  `openpyxl` to read Microsoft Excel files
3. `mako` the Mako template library

Typical column names in the Microsoft Excel file may be "email", "name", "attendance_mode", where the first two fields are string and the third is either a string or a boolean in case the logic in the Mako template is a simple Yes or No.