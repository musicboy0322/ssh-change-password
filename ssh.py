# import things
import configparser
import paramiko
import win32com.client as win32
from datetime import datetime, timedelta
import random
from datetime import date
from functions import popResultWindow, logInformation, writeCsv, rewriteIni, sendingEmail

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

# if password's character lower than 8, will show error and quit this procedure 
if len(newPassword) < 8 :
    popResultWindow('This password is shorter than\n 8 characters, please change it')
    quit()

# get config information(storing log name and location)
logFileName = config['LOG']['FILENAME']
logLocation = config['LOG']['LOCATION']

# get config information(storing csv name)
csvFileName = config['CSV']['FILENAME']

# get config information(sending email)
email = config['SEND']['EMAIL']
emailSendYesNo = config['SEND']['EMAILSENDYESNO']

# gahter outlook user inforamtion
outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# select emails in 24 hours and judge whether there is expire password or not
changeYesNo = False
brutalChangeYesNo = False
for num in range(len(mapi.Folders)) :
    received_dt = datetime.now() - timedelta(days = 30)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M')
    messages = mapi.Folders(num + 1).Folders('收件匣').Items
    messages = messages.Restrict("[ReceivedTime] >='" + received_dt + "'")
    for msg in list(messages):
        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) <= 0:
            changeYesNo = True
        '''
        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) < 0:
            brutalChangeYesNo = True
        '''

# 0 days
if changeYesNo :
    try :
        # create a transport instance
        connection = paramiko.SSHClient()

        # create connection and specify it in sshclient
        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        connection.connect(hostname=hostname, port=port, username=username, password=password)

        # excute command
        stdin, stdout, stderr = connection.exec_command('passwd')
        stdin.write(newPassword + '\n')
        stdin.write(newPassword + '\n')
        stdin.flush()
        stdout.channel.set_combine_stderr(True)
        print(stdout.read().decode())

        # close connection
        connection.close()

        # setting log file's detail, location, text
        logInformation(logLocation, logFileName, 'all authentication tokens updated successfully')

        # synchronize changing config file's password information
        rewriteIni(config, newPassword)
        
        # write change result into csvfile
        writeCsv(csvFileName, date.today(), hostname, port, username, newPassword)

        # pop result window
        popResultWindow("Password has changed")

        # sending email about the result
        if emailSendYesNo == 'yes':
            sendingEmail(f'Your new password is {newPassword}', email, outlook)

    except Exception as e:
        print(e)

# less than 0 days
elif brutalChangeYesNo :
    try :
        # create a sshclient instance
        connection = paramiko.SSHClient()

        # create connection and specify it in sshclient
        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        connection.connect(hostname=hostname, port=port, username=username, password=password)
        interact = connection.invoke_shell()

        # read information and send text
        buff = ''
        while not buff.endswith("UNIX password: '"):
            resp = interact.recv(9999)
            buff += str(resp)
        interact.send(password + '\n')
        
        buff = ''
        while not buff.endswith("New password: '"):
            resp = interact.recv(9999)
            buff += str(resp)
        interact.send(newPassword + '\n')        

        buff = ''
        while not buff.endswith("Retype new password: '"):
            resp = interact.recv(9999)
            buff += str(resp)
        interact.send(newPassword + '\n')
        resp = interact.recv(9999)

        # close connection
        connection.close()

        # setting log file's detail, location, text
        logInformation(logLocation, logFileName, 'all authentication tokens updated successfully')

        # synchronize changing config file's password information
        rewriteIni(config, newPassword)
        
        # write change result into csvfile
        writeCsv(csvFileName, date.today(), hostname, port, username, newPassword)
        
        # pop result window
        popResultWindow("Password has changed")

        # sending email about the result
        if emailSendYesNo == 'yes':
            sendingEmail(f'Your new password is {newPassword}', email, outlook)

    except Exception as e:
        print(e)

else:
    # setting log file's detail, location, text
    logInformation(logLocation, logFileName, 'Not yet to change')

    # pop result window
    popResultWindow('Not yet to change')

    # sending email about the result
    if emailSendYesNo == 'yes':
        sendingEmail('Not yet to change', email, outlook)



        

        

