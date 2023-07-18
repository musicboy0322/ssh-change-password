import configparser
import logging
import paramiko
import win32com.client as win32
from datetime import datetime, timedelta
import random
import tkinter as tk
import csv
import os 
from datetime import date

# import config file
config = configparser.ConfigParser()
config.read('config.ini', encoding = 'utf-8')

# get config information(target server)
username = config['TARGET']['USERNAME']
password = config['TARGET']['PASSWORD']
hostname = config['TARGET']['HOSTNAME']
port = int(config['TARGET']['PORT'])

# get config information(new password) and choose random password without repeating the same password
newPasswordSplit = config['NEW']['PASSWORD'].split(',')
newPassword = newPasswordSplit[random.randint(0,len(newPasswordSplit)-1)]
while newPassword == password :
    newPassword = newPasswordSplit[random.randint(0,len(newPasswordSplit)-1)]

# get config information(storing log name and location)
logFileName = config['LOG']['FILENAME']
logLocation = config['LOG']['LOCATION']

# get config information(storing csv name)
csvFileName = config['CSV']['FILENAME']

# get config information(sending email)
email = config['SEND']['EMAIL']

# gahter outlook user inforamtion
outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# select emails in 24 hours and judge whether there is expire password or not
changeYesNo = False
for num in range(len(mapi.Folders)) :
    received_dt = datetime.now() - timedelta(days = 1)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M')
    messages = mapi.Folders(num + 1).Folders('收件匣').Items
    messages = messages.Restrict("[ReceivedTime] >='" + received_dt + "'")
    for msg in list(messages):
        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) <= 0:
            changeYesNo = True

if changeYesNo :
    try :
        # create a transport instance
        trans = paramiko.Transport((hostname, port))

        # create connection and specify it in sshclient
        trans.connect(username=username, password=password)
        ssh = paramiko.SSHClient()
        ssh._transport = trans

        # excute command
        stdin, stdout, stderr = ssh.exec_command('passwd')
        stdin.write(newPassword + '\n')
        stdin.write(newPassword + '\n')
        stdin.flush()
        stdout.channel.set_combine_stderr(True)
        print(stdout.read().decode())

        # close connection
        trans.close()

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

        logging.info('all authentication tokens updated successfully.')

        # synchronize changing config file's password information
        config.set('TARGET', 'PASSWORD', newPassword)
        with open('config.ini', 'w') as configfile:
            config.write(configfile)
        
        # write change result into csvfile
        if os.path.exists(csvFileName):
            with open(csvFileName, 'a', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([date.today(), hostname, port, username, newPassword])
        else:
            with open(csvFileName, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Date', 'Hostname', 'Port', 'Username', 'NewPassword'])
                writer.writerow([date.today(), hostname, port, username, newPassword])
        
        # pop result window
        window = tk.Tk()
        window.title("Result")
        window.geometry('250x90')
        pop = tk.Label(window,text="Password has changed", font=("Arial", 12), width=20,height=10)
        pop.pack()
        window.mainloop()

    except Exception as e:
        print(e)
else:
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

    logging.info('Not yet to change')

    # pop result window
    window = tk.Tk()
    window.title("Result")
    window.geometry('250x90')
    pop = tk.Label(window,text="Not yet to change", font=("Arial", 12), width=20,height=10)
    pop.pack()
    window.mainloop()




# customize function
def sendingEmail():
    # sending email about the result of changing password
    mail = outlook.CreateItem(0)
    mail.Subject = 'Auto Changing Password Result'
    mail.Body = f'Your New Password : {newPassword}'
    mail.To = email
    mail.Send()
    print('Sending successful')