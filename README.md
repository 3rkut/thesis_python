# Thesis on web parsing automation with python
Python code for my thesis

# Libraries used
- json
- pandas
- smtplib
- requests
- email

# Where is the data comes from?
- Tefas.gov.tr

# Data format
1. JSON content
2. JSON parsing
3. Creating excel report(xlsx)

# Functionalities
- Stock prices analysis by months and years.
- Calculating highest and lowest stocks prices and information about the stock.
- With all this information, calculated prices and other analysis generates the excel report.
- When we first launched the python script, it asks for e-mail address to send excel report.
- For this reason, in lines 45-46, you need to edit your mail creds for auth to do that.(For sending mails to target e-mail with attachment of the excel file)
