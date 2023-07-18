import configparser
import logging
import paramiko
import win32com.client as win32
from datetime import datetime, timedelta
import random
import tkinter as tk

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

# get config information(storing name and location)
fileName = config['STORE']['FILENAME']
location = config['STORE']['LOCATION']

# get config information(sending email)
email = config['SEND']['EMAIL']

# gahter outlook user inforamtion
outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

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
        if len(location) == 0 :
            logging.basicConfig(
                filename = fileName,
                format = '%(asctime)s %(levelname)s %(message)s',
                level = logging.INFO
            )
        else :
            logging.basicConfig(
                filename = location + '/' + fileName,
                format = '%(asctime)s %(levelname)s %(message)s',
                level = logging.INFO
            )

        logging.info(f'Your new password : {newPassword}')

        # synchronize changing config file's password information
        config.set('TARGET', 'PASSWORD', newPassword)
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

        # pop result window
        window = tk.Tk()
        window.title("Result")
        window.geometry('250x90')
        l = tk.Label(window,text="Password has changed", font=("Arial", 12), width=20,height=10)
        l.pack()
        window.mainloop()

    except Exception as e:
        print(e)
else:
    # pop result window
    window = tk.Tk()
    window.title("Result")
    window.geometry('250x90')
    l = tk.Label(window,text="Not yet to change", font=("Arial", 12), width=20,height=10)
    l.pack()
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