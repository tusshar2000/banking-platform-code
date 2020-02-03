import openpyxl
openpyxl.__version__
from openpyxl import load_workbook
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random
import time

wbook=load_workbook("Workbook.xlsx")#name of the excel file and it should be present in same directory
                                    #or else you need to specifiy the whole path
                                    #use ctrl+F, find wherever "Workbook.xlxs" appears, make changes if you used
                                    #a different name.
sh1=wbook["UserDetails"] #name of sheet1

#OTP generating function
def OTP(x,y):

    fromaddr="rtu780002@gmail.com" #use your email here and it should be a string
    toaddr=y
    msg=MIMEMultipart()
    
    msg['Subject']="OTP generated for bank process"

    body="Below is your 4 digit OTP, use this for your bank process and don't share it with anyone else.\n%d"%(x)
    msg.attach(MIMEText(body,'plain'))

    server=smtplib.SMTP("smtp.gmail.com")
    server.starttls()
    #tls-->transport layer security
    server.login("rtu780002@gmail.com","xyz") #in place of xyz enter password of your email used
    text=msg.as_string()
    server.sendmail("fromaddr",toaddr,text) #couldnt use toaddr so had to write the email
    server.quit()

flag=0
d1={}
for i in range(4):
    #This "if" code runs when you input wrong name entries thrice
    if i==3:
        print("You made three wrong entries, quitting this session.")
        flag=1
        break
        
    elif flag!=1:
        Name=input("Enter your full name or UserId:")
        for i in range(2,13):
            if Name==sh1['D%d'%(i)].value or Name==str(sh1['B%d'%(i)].value):
                for j in range(4):
                    #This if code runs when you input wrong 
                    #account number of the corresponding name thrice 
                    if j==3:
                        print("You made three wrong entries, changing account status to blocked")
                        print('Quit Program')
                        sh1['G%d'%(i)].value='Blocked'
                        wbook.save("Workbook.xlsx")
                        flag=1
                        break
                            
                    elif flag!=1:
                        bankaccno=input("Enter your bank account no:")
                    
                        if bankaccno==sh1['C%d'%(i)].value:
                            print("Okay")
                    
                            y=sh1['E%d'%(i)].value  # y contains the email address of the person.
                            x= random.randint(1000,9999)# x holds a randomly generated 4 digit number.

                            #we pass x and y as parameters in OTP function which mails you the OTP.
                            OTP(x,y)
                            
                            
                            for k in range (4):
                                #This if runs when you input wrong OTP thrice.
                                if k==3:
                                    print("You made three wrong entries, changing account status to blocked")
                                    print('Quit Program')
                                    sh1['G%d'%(i)].value='Blocked'
                                    wbook.save("Workbook.xlsx")
                                    flag=1
                                    break
                                elif flag!=1:
                                    OTP=int(input("Enter OTP: "))
                        
                                    if OTP==x:
                                        print("Okay")
                                        
                                        for t in range(4):
                                            
                                            #If user makes any other entries other than 'C' 
                                            #(refers credit) or 'D'(refers debit) thrice, account gets
                                            #blocked.
                                            if t==3:
                                                print("You made three wrong entries, changing account status to blocked")
                                                print('Quit Program')
                                                sh1['G%d'%(i)].value='Blocked'
                                                wbook.save("Workbook.xlsx")
                                                flag=1
                                                break
                                            elif flag!=1:
                                                typeoftrans=input("Enter 'C' for credit and 'D' for Debit")
                                                if typeoftrans=='C':
                                                    d1={'initial':sh1['F%d'%(i)].value}
                                                    money=int(input("Enter Amount to be credited."))
                                                    sh1['F%d'%(i)].value+=money
                                                    localtime = time.asctime( time.localtime(time.time()) )
                                                    s7="credit"
                                                    d1[s7]=money
                                                    d1['balance']=sh1['F%d'%(i)].value
                                                    flag=1
                                                    wbook.save("Workbook.xlsx")
                                                    break
                                            
                                                elif typeoftrans=='D':
                                                    d1={'initial':sh1['F%d'%(i)].value}
                                                    money=int(input("Enter Amount to be credited."))
                                                    sh1['F%d'%(i)].value-=money
                                                    localtime = time.asctime( time.localtime(time.time()) )
                                                    s7="debit"
                                                    d1[s7]=money
                                                    d1['balance']=sh1['F%d'%(i)].value
                                                    flag=1
                                                    wbook.save("Workbook.xlsx")
                                                    break
                                            
                                                else:
                                                    print("Wrong character.\n")
                                            elif flag==1:
                                                break
                                    
                                    else:
                                        print("Wrong entry\n")
                                elif flag==1:
                                        break
                        else:
                            print("Wrong bank account number\n")
                    elif flag==1:
                        break
        else:
            if flag!=1:
                print("Sorry no such name exists in our database.\n")     
            else:
                break
    elif flag==1:
        break

#Prints the mini statement.
if len(d1)>1:
    print("\n****Mini Statement****")
    print(Name)
    print(localtime)
    for i in d1:
        print(i,d1[i])