import datetime as dt
import os
import win32com.client

def run():
    today = dt.datetime.today()
    year = today.strftime('%Y')
    month = today.strftime('%m-%B')

    src_path = 'M:/FPPShare/FPP-Production/NCOA Files'

    path = f'{src_path}/{year}/{month}/'

    if os.path.exists(path) and len(os.listdir(path)) > 0:
        files = [file for file in os.listdir(path) if not 'done' in file and not '~' in file]
        
    files = '\n'.join(files)

    outlook = win32com.client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = "DEnglish2@northwell.edu, dpashayan@northwell.edu"
    message.Body = f"There is currently a file in the NCOA folder for the bot to work. The file names are: \n {files}"
    message.Subject = "NCOA - Files for Bot"
    message.Send()
    
if __name__ == '__main__':
    run()