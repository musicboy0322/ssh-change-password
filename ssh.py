# import things
import json 
import paramiko
import win32com.client as win32
from datetime import datetime, timedelta
import random
from datetime import date
import logging
from functions import writeCsv, rewriteJson, sendingEmail, traverseFolders, displayProgressBar, generateRandomPassword
from progress.bar import ShadyBar 

try:

    # generate progress bar
    with ShadyBar('Reading config file', max=100, suffix='%(percent)d%%') as bar:

        # import config file
        with open('config.json', 'r') as configFile:
            config = json.load(configFile)

        displayProgressBar(bar, 100)
        bar.finish()

    # get config information(storing log name and location)
    logFileName = config['LOG']['filename']
    logLocation = config['LOG']['location']

    # get config information(sending email)
    email = config['SEND']['email']
    emailSendYesNo = config['SEND']['emailSendYesNo']

    # get config information(storing csv name)
    csvFileName = config['CSV']['filename']

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

    # gahter outlook user inforamtion
    outlook = win32.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    # select emails in 24 hours and judge whether there is expire password or not
    targetMail = []

    # generate progress bar
    with ShadyBar('Searching target email', max=100, suffix='%(percent)d%%') as bar:

        for i in range(len(mapi.Folders)) :
            root_folder = mapi.Folders[i].Folders[1]
            category = traverseFolders(root_folder, datetime, timedelta, targetMail)

        # generate progress bar(progress bar is accroding to how many top folders you have and will calculate every time it will increase)
        displayProgressBar(bar, (100 / len(mapi.Folders)) * (i + 1))
        bar.finish()

    if len(targetMail) == 0 :
        # writing log text
        logging.info('Password not yet to change')

        # print result
        print('Password not yet to change\n')

        # sending email about the result
        if emailSendYesNo == 'yes':
            sendingEmail('Password not yet to change', email, outlook)

        # as the sentence say
        input('Press Enter to exit...')

        exit()
    
    logging.info(f'you have {len(targetMail)} server need to change password')

    # this loop is to traverse target email
    for k in range(len(targetMail)):
        # category
        category = targetMail[k][1]

        # target host name
        targetHostName = targetMail[k][0]

        # this loop is to find
        for c in range(len(config['TARGET'])):
            if targetHostName == config['TARGET'][c]['hostname']:
                # get config information(target server)
                username = config['TARGET'][c]['username']
                password = config['TARGET'][c]['password']
                hostname = config['TARGET'][c]['hostname']
                port = config['TARGET'][c]['port'] 

                # as it say, generate random new password
                newPassword = generateRandomPassword()
                
                # 0 days
                if category == 'change' :
                    # writing log text
                    logging.info('password will expire in 0 day, password need to change')

                    # generate progress bar
                    with ShadyBar('Connecting server', max=100, suffix='%(percent)d%%') as bar:

                        # create sshclient instance
                        connection = paramiko.SSHClient()

                        # generate progress bar
                        displayProgressBar(bar, 50)

                        # create connection and specify it in sshclient
                        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                        connection.connect(hostname=hostname, port=port, username=username, password=password)

                        # generate progress bar
                        displayProgressBar(bar, 100)
                        bar.finish()

                    # writing log text
                    logging.info('Successfully connect to remote server')

                    # generate progress bar
                    with ShadyBar('Changing password', max=100, suffix='%(percent)d%%') as bar:

                        # excute command
                        stdin, stdout, stderr = connection.exec_command('passwd')
                        stdin.write(password + '\n')
                        stdin.write(newPassword + '\n')
                        stdin.write(newPassword + '\n')
                        stdout.channel.set_combine_stderr(True)
                        stdout.read().decode()

                        # close connection
                        connection.close()

                        # generate progress bar
                        displayProgressBar(bar, 100)
                        bar.finish()

                    # writing log text
                    logging.info('All authentication tokens updated successfully')

                    # synchronize changing config file's password information
                    rewriteJson(config, newPassword, c)

                    # generate progress bar
                    with ShadyBar('Writing csv', max=100, suffix='%(percent)d%%') as bar:

                        # write change result into csvfile
                        writeCsv(csvFileName, date.today(), hostname, port, username, newPassword)

                        # generate progress bar
                        displayProgressBar(bar, 100)
                        bar.finish()

                    # sending email about the result
                    if emailSendYesNo == 'yes':
                        sendingEmail(f'Your new password is {newPassword}', email, outlook)

                # less than 0 days
                elif category == 'brutal' :
                    # writing log text
                    logging.info('Password will expire in less than 0 days, password need to change')

                    # generate progress bar
                    with ShadyBar('Connecting server', max=100, suffix='%(percent)d%%') as bar:
                    
                        # create a sshclient instance
                        connection = paramiko.SSHClient()

                        # generate progress bar
                        displayProgressBar(bar, 50)

                        # create connection and specify it in sshclient
                        connection.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                        connection.connect(hostname=hostname, port=port, username=username, password=password)
                        interact = connection.invoke_shell()

                        # generate progress bar
                        displayProgressBar(bar, 100)
                        bar.finish()

                    # writing log text
                    logging.info('Successfully connect to remote server')

                    # generate progress bar
                    with ShadyBar('Changing password', max=100, suffix='%(percent)d%%') as bar:

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

                        # close connection
                        connection.close()

                        # generate progress bar
                        displayProgressBar(bar, 100)
                        bar.finish()

                    # writing log text
                    logging.info('All authentication tokens updated successfully')

                    # synchronize changing config file's password information
                    rewriteJson(config, newPassword, c)
                    
                    # generate progress bar
                    with ShadyBar('Writing csv', max=100, suffix='%(percent)d%%') as bar:

                        # write change result into csvfile
                        writeCsv(csvFileName, date.today(), hostname, port, username, newPassword)

                        # generate progress bar
                        displayProgressBar(bar, 100)
                        bar.finish()

                    # sending email about the result
                    if emailSendYesNo == 'yes':
                        sendingEmail(f'Your new password is {newPassword}', email, outlook)

    print('Successfully change all the password\n')

    # as the sentence say
    input('Press Enter to exit...')            
                
except Exception as e:
    print(e)

    # as the sentence say
    input('Press Enter to exit...')