import tkinter as tk
import logging
import csv
import os 

# used function
def popResultWindow(text):
    # as the function name, pop result window
    window = tk.Tk()
    window.title("Result")
    window.geometry('250x90')
    pop = tk.Label(window,text=text, font=("Arial", 12), width=20,height=10)
    pop.pack()
    window.mainloop()

def logInformation(logLocation, logFileName, text):
    # setting log file's detail and location
    if len(logLocation) == 0 :
        logging.basicConfig(
            filename = logFileName,
            format = '%(asctime)s %(levelname)s %(message)s',
            level = logging.INFO
        )
    else :
        logging.basicConfig(
            filename = logLocation + '/' + logFileName,
            format = '%(asctime)s %(levelname)s %(message)s',
            level = logging.INFO
        )

    logging.info(text)

def writeCsv(csvFileName, date, hostname, port, username, newPassword) :
    # write change result into csvfile
    if os.path.exists(csvFileName):
        with open(csvFileName, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([date.today(), hostname, port, username, newPassword])
    else:
        with open(csvFileName, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Date', 'Hostname', 'Port', 'Username', 'NewPassword'])
            writer.writerow([date, hostname, port, username, newPassword])

def rewriteIni(config, newPassword):
    # synchronize changing config file's password information
    config.set('TARGET', 'PASSWORD', newPassword)
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

def sendingEmail(text, email, outlook):
    # sending email about the result of changing password
    mail = outlook.CreateItem(0)
    mail.Subject = 'Auto Changing Password Result'
    mail.Body = text
    mail.To = email
    mail.Send()
    print('Sending successful')