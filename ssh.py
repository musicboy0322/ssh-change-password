# import things
import configparser
import paramiko
import win32com.client as win32
from datetime import datetime, timedelta
import random
from datetime import date
import logging
from functions import popResultWindow, writeCsv, rewriteIni, sendingEmail, traverseFolders

# import config file
config = configparser.ConfigParser()
config.read('config.ini', encoding = 'utf-8')

# get config information(storing log name and location)
logFileName = config['LOG']['FILENAME']
logLocation = config['LOG']['LOCATION']

# get config information(storing csv name)
csvFileName = config['CSV']['FILENAME']

# get config information(sending email)
email = config['SEND']['EMAIL']
emailSendYesNo = config['SEND']['EMAILSENDYESNO']

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

# if password's character lower than 8, will show error and quit this procedure 
if len(newPassword) < 8 :
    popResultWindow('This password is shorter than\n 8 characters, please change it')

    # writing log text
    logging.warning('Bad password, the password is shorter than 8 characters')

    # stop the whole procedure
    quit()

# gahter outlook user inforamtion
outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# select emails in 24 hours and judge whether there is expire password or not
for i in range(len(mapi.Folders)) :
    root_folder = mapi.Folders[i].Folders[1]
    category = traverseFolders(root_folder, datetime, timedelta)

# 0 days
if category == 'change' :
    try :
        # writing log text
        logging.info('password will expire in 0 day, password need to change')

        # create sshclient instance
        connection = paramiko.SSHClient()

        # create connection and specify it in sshclient
        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        connection.connect(hostname=hostname, port=port, username=username, password=password)

        # writing log text
        logging.info('Successfully connect to remote server')

        # excute command
        stdin, stdout, stderr = connection.exec_command('passwd')
        stdin.write(newPassword + '\n')
        stdin.write(newPassword + '\n')
        stdin.flush()
        stdout.channel.set_combine_stderr(True)

        # writing log text
        logging.info(stdout.read().decode())

        # close connection
        connection.close()

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
        logging.error(e)

# less than 0 days
elif category == 'brutal' :
    try :
        # writing log text
        logging.info('Password will expire in less than 0 days, password need to change')

        # create a sshclient instance
        connection = paramiko.SSHClient()

        # create connection and specify it in sshclient
        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        connection.connect(hostname=hostname, port=port, username=username, password=password)
        interact = connection.invoke_shell()

        # writing log text
        logging.info('Successfully connect to remote server')

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

        # writing log text
        logging.info('All authentication tokens updated successfully')

        # close connection
        connection.close()

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
        logging.error(e)

else:
    # writing log text
    logging.info('Password not yet to change')

    # pop result window
    popResultWindow('Password not yet to change')

    # sending email about the result
    if emailSendYesNo == 'yes':
        sendingEmail('Password not yet to change', email, outlook)