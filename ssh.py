# import things
import configparser
import paramiko
import win32com.client as win32
from datetime import datetime, timedelta
import random
from datetime import date
import logging
from functions import writeCsv, rewriteIni, sendingEmail, traverseFolders, displayProgressBar

# generate progress bar
displayProgressBar('Reading config', 0)

# import config file
config = configparser.ConfigParser()
config.read('config.ini', encoding = 'utf-8')

# generate progress bar
displayProgressBar('Reading config', 100)

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

    # print result
    print('This password is shorter than\n 8 characters, please change it\n')

    # writing log text
    logging.warning('Bad password, the password is shorter than 8 characters')

    # as the sentence say
    input('Press Enter to exit...')

    # stop the whole procedure
    quit()

# gahter outlook user inforamtion
outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# generate progress bar
displayProgressBar('Searching target email', 0)

# select emails in 24 hours and judge whether there is expire password or not
for i in range(len(mapi.Folders)) :
    root_folder = mapi.Folders[i].Folders[1]
    category = traverseFolders(root_folder, datetime, timedelta)

    # generate progress bar(progress bar is accroding to how many top folders you have and will calculate every time it will increase)
    displayProgressBar('Searching target email', (100 / len(mapi.Folders)) * (i + 1))

# 0 days
if category == 'change' :
    try :
        # writing log text
        logging.info('password will expire in 0 day, password need to change')

        # generate progress bar
        displayProgressBar('Connecting server', 0)

        # create sshclient instance
        connection = paramiko.SSHClient()

        # generate progress bar
        displayProgressBar('Connecting server', 50)

        # create connection and specify it in sshclient
        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        connection.connect(hostname=hostname, port=port, username=username, password=password)

        # generate progress bar
        displayProgressBar('Connecting server', 100)

        # writing log text
        logging.info('Successfully connect to remote server')

        # generate progress bar
        displayProgressBar('Changing password', 0)

        # excute command
        stdin, stdout, stderr = connection.exec_command('passwd')
        stdin.write(newPassword + '\n')
        stdin.write(newPassword + '\n')
        stdin.flush()
        stdout.channel.set_combine_stderr(True)

        # generate progress bar
        displayProgressBar('Changing password', 100)

        # writing log text
        logging.info(stdout.read().decode())

        # close connection
        connection.close()

        # synchronize changing config file's password information
        rewriteIni(config, newPassword)

        # generate progress bar
        displayProgressBar('Writing csv', 0)

        # write change result into csvfile
        writeCsv(csvFileName, date.today(), hostname, port, username, newPassword)

        # generate progress bar
        displayProgressBar('Writing csv', 100)

        # print result
        print("Password has changed\n")

        # sending email about the result
        if emailSendYesNo == 'yes':
            sendingEmail(f'Your new password is {newPassword}', email, outlook)
        
        # as the sentence say
        input('Press Enter to exit...')
        

    except Exception as e:
        logging.error(e)

# less than 0 days
elif category == 'brutal' :
    try :
        # writing log text
        logging.info('Password will expire in less than 0 days, password need to change')

        # generate progress bar
        displayProgressBar('Connecting server', 0)
        
        # create a sshclient instance
        connection = paramiko.SSHClient()

        # generate progress bar
        displayProgressBar('Connecting server', 50)

        # create connection and specify it in sshclient
        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        connection.connect(hostname=hostname, port=port, username=username, password=password)
        interact = connection.invoke_shell()

        # generate progress bar
        displayProgressBar('Connecting server', 100)

        # writing log text
        logging.info('Successfully connect to remote server')

        # generate progress bar
        displayProgressBar('Changing password', 0)

        # read information and send text
        buff = ''
        while not buff.endswith("password: '"):
            resp = interact.recv(9999)
            buff += str(resp)
        interact.send(password + '\n')
        
        buff = ''
        while not buff.endswith("password: '"):
            resp = interact.recv(9999)
            buff += str(resp)
        interact.send(newPassword + '\n')        

        buff = ''
        while not buff.endswith("password: '"):
            resp = interact.recv(9999)
            buff += str(resp)
        interact.send(newPassword + '\n')
        resp = interact.recv(9999)
        resp = interact.recv(9999)

        # generate progress bar
        displayProgressBar('Changing password', 100)

        # writing log text
        logging.info('All authentication tokens updated successfully')

        # close connection
        connection.close()

        # synchronize changing config file's password information
        rewriteIni(config, newPassword)
        
        # generate progress bar
        displayProgressBar('Writing csv', 0)

        # write change result into csvfile
        writeCsv(csvFileName, date.today(), hostname, port, username, newPassword)

        # generate progress bar
        displayProgressBar('Writing csv', 100)
        
        # print result
        print("Password has changed\n")

        # sending email about the result
        if emailSendYesNo == 'yes':
            sendingEmail(f'Your new password is {newPassword}', email, outlook)

        # as the sentence say
        input('Press Enter to exit...')

    except Exception as e:
        logging.error(e)

else:
    # writing log text
    logging.info('Password not yet to change')

    # print result
    print('Password not yet to change\n')

    # sending email about the result
    if emailSendYesNo == 'yes':
        sendingEmail('Password not yet to change', email, outlook)

    # as the sentence say
        input('Press Enter to exit...')