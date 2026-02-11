# `bulk_email`

`bulk_email` send emails using Gmail SMTP to a list of recipients whose name, email address and other parameters are read from a Microsoft Excel file.

It uses Mako template library to implement processing logic and looping. Data of each recipient is stored as a tuple with the first item is the email address and the second is the name of the recipient. The rest can be any other data required by the template.