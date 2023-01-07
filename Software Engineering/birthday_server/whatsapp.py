import pywhatkit #for whats app
from datetime import datetime as dt
import datetime
import senders_data #email and password of sender from another file
import smtplib  #simple message transfer protocol#library to send email to users email_address
import pandas
def whatsapp_msg(name, phone, msg):
    # Set time to send msg i.e 1 min later than current time
    hr = int(datetime.datetime.now().hour)
    min = int(datetime.datetime.now().minute) + 1

    pywhatkit.sendwhatmsg(phone, msg, hr, min)
    # WhatsApp will open and after 15 Seconds  & message will be Delivered!

    print("Messsage sent to " + str(name) +
          " wishing them with message: " + str(msg))



server =smtplib.SMTP('smtp.gmail.com',587)
Senders_email = senders_data.email
Senders_password= senders_data.password
def login_into_sendersemail():
    server.starttls()   #transfered layer security
    server.login(Senders_email,password=Senders_password)
# receivers_name=input("Enter receivers name ")
# receivers_email=input("Enter receivers email ")

dataframe = pandas.read_excel("New Microsoft Excel Worksheet.xlsx")
today = datetime.datetime.now().strftime("%d-%m")
yearNow = datetime.datetime.now().strftime("%Y")

for index, item in dataframe.iterrows():
    # stripping the birthdays in excel sheet as : DD-MM

    bday = item['Birthday'].strftime("%d-%m")

    receivers_name = item['Name']
    phone = '+91' + str(item['Phone'])
    receivers_email = item['Email']

    # Message to be sent
    msg = "Wish you a very happy birthday " + receivers_name + "! \nFrom Aarush"


def send_email():

    login_into_sendersemail()
    msg=("Hi "+ receivers_name +","+"\n"+"Happy birthday ")
    print(msg)
    server.sendmail(Senders_email,receivers_email,msg)
    server.quit() #to quit the server
    print("greeting has been sent successfully!")

#send emailfunction called
if __name__ == "__main__":
    # send_email()
    # whatsapp_msg("sangam", "happy birthday sangam")
    if today==bday:
        send_email()
    else:
        print("nobody in the database has birthday today!")
    if today==bday:
        whatsapp_msg("sangam",phone, "happy birthday sangam")
    else:
        print("nobody in the database has birthday today!")



