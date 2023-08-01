import tkinter as tk
import csv
import os
from progress.bar import ShadyBar 

# used function
def popResultWindow(text):
    # as the function name, pop result window
    window = tk.Tk()
    window.title("Result")
    window.geometry('280x90')
    pop = tk.Label(window,text=text, font=("Arial", 12), width=30,height=10)
    pop.pack()
    window.mainloop()

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

def traverseFolders(folder, datetime, timedelta):
    messages = folder.Items
    received_dt = datetime.now() - timedelta(days = 1)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt +"'")
    message_count = messages.Count
    for i in range(1, message_count + 1):
        msg = messages.Item(i)
        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) == 0:
            return 'change'

        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) < 0:
            return 'brutal'

    subfolders = folder.Folders
    for subfolder in subfolders:
        result = traverseFolders(subfolder, datetime, timedelta)
        if result == 'brutal':
            return 'brutal'

        elif result == 'change':
            return 'change'

def displayProgressBar(barName, degree):
    # generate progress bar
    with ShadyBar(barName, max=100, suffix='%(percent)d%%') as bar:
        bar.next(degree)
        bar.finish()