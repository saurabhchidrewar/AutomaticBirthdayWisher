'''

AUTOMATIC BIRTHDAY REMINDER USING PYTHON

PROJECT DESIGNED BY-
SAURABH CHIDREWAR

'''

#IMPORTING REQUIRED LIBRARIES
import pandas as pd
import datetime
import smtplib
import os

#CHANGING DIRECTORY (FOR WINDOWS TASK SCHEDULER)
os.chdir(r"#") #PROJECT DIRECTORY PATH

#AUTHENTICATION
GMAIL_ID = '#'
GMAIL_PSWD = '#'

#SEND EMAIL FUNCTION USING smtplib 
def sendEmail(to,sub,message):
    print(f"Email to {to} sent with subject: {sub} and message {message}")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID,to,f"Subject: {sub}\n\n{message}")
    s.quit()

#MAIN FUNCTION 
if (__name__=="__main__"):
    #READING LOCAL EXCEL FILE CONTAINING DATA
    df = pd.read_excel("data.xlsx")
    
    #READING CURRENT DATE,MONTH AND YEAR
    today = datetime.datetime.now().strftime("%d-%m")
    currentYear = datetime.datetime.now().strftime("%Y")

    writeInd = [] #This list will store the indices of the people having their birthday

    for index,item in df.iterrows():
        #STORE BIRTHDAY OF THE CURRENT MEMBER
        bday = item['Birthday'].strftime('%d-%m') 

        #COMPARE THE BIRTHDAY OF THE MEMBER WITH THE CURRENT DATE
        if (today==bday and currentYear not in df['Year']):
            sendEmail(item['Email'],"Happy Birthday",item['Dialogue'])

            #APPEND THE INDICES LIST 
            writeInd.append(index)

    if (len(writeInd)):
        #ADD CURRENT YEAR TO THE DATA SHEET 
        '''
        This will help in determining that the individual was wished this year
        ''' 
        for i in writeInd:
            #MODIFY THE YEAR COLUMN OF THE DATAFRAME
            yr = df.loc[i,'Year']
            df.loc[i,'Year'] = str(yr) + ', ' + str(currentYear)

        #RENDER THE DATAFRAME INTO AN EXCEL FILE
        df.to_excel('data.xlsx',index=False)

'''
Excel Data File Contains:
Sno	Name	Birthday	Dialogue	Year	Email
as columns 
'''