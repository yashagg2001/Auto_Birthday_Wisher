# Auto_Birthday_Wisher
I created a python script that will automatically deliver birthday or anniversary messages to my friends or relatives on their special occasion.<br>

`Firstly update and add the important dates in data.xlsx file.`<br>
# Libraries used
1. notification from plyer: `for showing pop up reminder notifications`
2. smtplib: `for sending emails`
3. pywhatkit: `for sending whatsapp messages`
4. <b>other inbuilt libraries need to be imported</b>: `pandas, numpy, datetime, time, os`

# How it works
1. Using Task scheduler the python script `bday.py` run automatically every midnight and check if there is any important date. 
2. If it is so, then a <b>`notification reminder`</b> will pop up showing that to whom it will be going to wish. 
3. After that, whatsapp web will open and message will be sent to that person. 
4. If email id of that person is also present in the data excel sheet then a wish message is also sent via an email.

