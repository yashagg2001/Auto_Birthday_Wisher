from numpy.core.numeric import NaN
import pandas as pd
import datetime
import smtplib
import os
os.chdir(r"D:\Projects\Birthday_Wisher")

def notifier(name,today,msg,status):
    import time
    from plyer import notification
    notification.notify(
        title= f"****Happy {status} {name}****",
        message=f"Today is {str(today)}.\nMake sure that your whatsapp web is logged in on your PC.\nThe message that will be sent on his/her whatsapp is:-\n{msg}",
        app_icon ="D:\Projects\Birthday_Wisher\cake.ico",
        timeout= 18
    )
    time.sleep(18)

#For sending birthday wish by Email using Gmail
def sendEmail(to, sub, msg):
    # Enter your GMail credentials
    Gmail_Email = ''
    Gmail_password = ''
    
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(Gmail_Email, Gmail_password)
    s.sendmail(Gmail_Email, to, f"Subject: {sub}\n\n{msg}")
    s.quit()
    print(f"Email to {to} sent successfully\nwith subject: {sub} and message is: {msg}\n" )

#For sending birthday wish using Whatsapp
'''
def sendWhatsappMsg(name,msg):
    import webbrowser
    import pyautogui
    import time

    webbrowser.open("https://web.whatsapp.com/")
    time.sleep(35)
    # pyautogui.click(138,392)
    pyautogui.click(170,257)
    pyautogui.typewrite(name)
    time.sleep(3)
    pyautogui.click(180,448)
    time.sleep(2)
    pyautogui.click(873,980)
    pyautogui.typewrite(msg)
    time.sleep(2)
    pyautogui.click(1864,983)
'''
def sendWhatsappMsg(mob,msg):
    import pywhatkit
    from plyer import notification
    try:
        mobile=f'+91{str(mob)}'[0:13]
        pywhatkit.sendwhatmsg(mobile,f'{str(msg)}',00,2)
        print('Message sent successfully')
        notification.notify(
        title= f"****Whatsapp message sent successfully****",
        message=f"Congrats, message has been sent.",
        timeout= 10)
    except:
        print("An unexcepted error occured!")
        notification.notify(
        title= f"****An unexpected error occured****",
        message=f"Message hasn't been sent",
        timeout= 10)


if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    
    writeInd = []
    for index, item in df.iterrows():
        bday = item['DOB'].strftime("%d-%m")
        if((today == bday) and yearNow not in str(item['Year'])):
            notifier(item['Name'],today,item['Message'],item['Status'])
            sendWhatsappMsg(item['Mobile'],item['Message'])
            # Set Email and Password above, before uncommenting this part
            '''if(str(item['Email'])!="nan"):
                sendEmail(item['Email'], f"Happy {item['Status']} {item['Name']}", item['Message'])''' 
            writeInd.append(index)

    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ', ' + str(yearNow)

    df.to_excel('data.xlsx', index=False)   