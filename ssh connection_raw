import re
import paramiko
import win32com.client as win32

def sendemail(emailto, emailsubject, emailbody = None):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = emailto
    mail.Subject = emailsubject
    mail.Body = emailbody
    mail.Send()

def ssh_files_el():

### SSH connection ###

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname='<ssh_host>', username='<username>', password='<password>')
    ftp = ssh.open_sftp()

### Download files with a certain pattern ###

    ssh_folder = '<ssh_folder_path>'
    local_folder = '<output_folder_path>'
    pattern = re.compile('<pattern>')

    for ssh_file in ftp.listdir(ssh_folder):
        if pattern.match(ssh_file):
            ftp.get(ssh_folder + ssh_file, local_folder + ssh_file)

    ftp.close()
    ssh.close()

### Send email if succeed ###

    sendemail('<email_addresses>', '<email_title>', '<email_body>')

###  Execution, send a different email if fail ###

try:
    ssh_files_el()
except Exception as error:
    sendemail('<email_addresses>', '<email_title>', 'Reason of Failure:' + repr(error))
