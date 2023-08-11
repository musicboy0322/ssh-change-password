import tkinter as tk
import csv
import os
from progress.bar import ShadyBar 
import secrets
import string
import random
import json 
import time

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

def rewriteJson(config, newPassword, number):
    # synchronize changing config file's password information
    config['TARGET'][number]['password'] = newPassword

    with open('config.json', 'w') as configFile:
        json.dump(config, configFile, indent=1)
        

def sendingEmail(text, email, outlook):
    # sending email about the result of changing password
    mail = outlook.CreateItem(0)
    mail.Subject = 'Auto Changing Password Result'
    mail.Body = text
    mail.To = email
    mail.Send()
    print('Sending successful')

def traverseFolders(folder, datetime, timedelta, targetMail):
    messages = folder.Items
    received_dt = datetime.now() - timedelta(days = 1)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt +"'")
    message_count = messages.Count
    for i in range(1, message_count + 1):
        msg = messages.Item(i)
        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) == 0:
            # put text in [0] and put category in [1]
            temp = []
            # temporary
            # temp.append(str(msg).split(' ')[1])

            # official
            temp.append(msg.body.split(':')[2].split(' ')[1].split('\r')[0])

            temp.append('change')
            targetMail.append(temp)

        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) < 0:
            # put text in [0] and put category in [1]
            temp = []

            # temporary
            # temp.append(str(msg).split(' ')[1])

            # official
            temp.append(msg.body.split(':')[2].split(' ')[1].split('\r')[0])
            
            temp.append('brutal')
            targetMail.append(temp)

    subfolders = folder.Folders
    for subfolder in subfolders:
        result = traverseFolders(subfolder, datetime, timedelta, targetMail)

def displayProgressBar(bar, degree):
    bar.next(degree)

'''
def displayProgressBar(barName, degree):
    # generate progress bar
    with ShadyBar(barName, max=100, suffix='%(percent)d%%') as bar:
        for i in range(100) :
            bar.next(i)
'''

def generateRandomPassword():
    # generate random password
    characters = string.ascii_letters + string.digits
    punctuations = string.punctuation.replace('\\','')
    punctuations = punctuations.replace('"','')

    password = [secrets.choice(string.ascii_uppercase),
                secrets.choice(string.ascii_lowercase),
                secrets.choice(string.digits),
                secrets.choice(punctuations),
                secrets.choice(punctuations)]

    for i in range(5):
        password.extend(secrets.choice(characters))
    
    random.shuffle(password)

    result = ''.join(password)

    return result