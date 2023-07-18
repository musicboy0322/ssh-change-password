import configparser
import logging
import paramiko
import win32com.client as win32
from datetime import datetime, timedelta
import re

# import config file
config = configparser.ConfigParser()
config.read('config.ini', encoding = 'utf-8')

# get config information(target server)
username = config['TARGET']['USERNAME']
password = config['TARGET']['PASSWORD']
hostname = config['TARGET']['HOSTNAME']
port = int(config['TARGET']['PORT'])

# get config information(new password)
newPassword = config['NEW']['PASSWORD']

# get config information(storing name and location)
fileName = config['STORE']['FILENAME']
location = config['STORE']['LOCATION']

outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

result = []
for num in range(len(mapi.Folders)) :
    received_dt = datetime.now() - timedelta(days = 1)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M')
    messages = mapi.Folders(num + 1).Folders('收件匣').Items
    messages = messages.Restrict("[ReceivedTime] >='" + received_dt + "'")
    for msg in list(messages):
        if 'WARNING' in str(msg) and int(str(msg).split(' ')[8]) <= 0:
            result.append(str(msg))

if len(result) > 0 :
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

        logging.info(f'Your new password is {newPassword}')

    except Exception as e:
        print(e)
else:
    print('No yet to change')

