import openpyxl
openpyxl.__version__
from openpyxl import load_workbook
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random
import time
#name of the excel file should be same or else change the name in every location
#and it should be present in same directory or else you need to specifiy the whole path.

wbook = load_workbook("Workbook.xlsx")
sh1 = wbook["UserDetails"] #name of sheet1

#OTP generating function
def OTP(number, to_address):
    #please make sure that in your email settings you enable access to less secured 
    #devices. Check your settings with the link provided.
    #https://myaccount.google.com/lesssecureapps

    msg = MIMEMultipart()
    msg['From'] = YOUR_EMAIL_GOES_HERE
    msg['Subject'] = "OTP generated for bank process"
    msg['To'] = to_address
    
    body = "Below is your 4 digit OTP, use this for your bank process and don't share it with anyone else.\n%d"%(x)
    msg.attach(MIMEText(body,'plain'))

    #tls-->transport layer security
    s = smtplib.SMTP(host = 'smtp.gmail.com', port = 587)
    s.starttls()
    s.login(YOUR_EMAIL_GOES_HERE, YOUR_EMAIL_PASSWORD_GOES_HERE) #in place of xyz enter password of your email used
    text = msg.as_string()
    s.sendmail("fromaddr", to_address, text) #couldnt use toaddr so had to write the email
    s.quit()

flag = 0
d1 = {}
for i in range(4):
    #This "if" code runs when you input wrong name entries thrice
    if i == 3:
        print("You made three wrong entries, quitting this session.")
        flag = 1
        break
        
    elif flag !=1:
        Name = input("Enter your full name or UserId: ")
        for i in range(2,13):
            if Name == sh1['D%d'%(i)].value or Name == str(sh1['B%d'%(i)].value):
                for j in range(4):
                    #This if code runs when you input wrong 
                    #account number of the corresponding name thrice 
                    if j == 3:
                        print("You made three wrong entries, changing account status to blocked")
                        print('Quit Program')
                        sh1['G%d'%(i)].value = 'Blocked'
                        wbook.save("Workbook.xlsx")
                        flag = 1
                        break
                            
                    elif flag != 1:
                        bankaccno = input("Enter your bank account no: ")
                    
                        if bankaccno == sh1['C%d'%(i)].value:
                            if sh1['G%d'%(i)].value == 'Blocked':
                                print("Please contact bank manager, activity for your has been blocked.")
                                flag = 1
                                break

                            print("Please wait for next step to load.")
                    
                            y = sh1['E%d'%(i)].value  # y contains the email address of the person.
                            x = random.randint(1000,9999)# x holds a randomly generated 4 digit number.
                            #we pass x and y as parameters in OTP function which mails you the OTP.
                            OTP(x,y)
                            
                            for k in range(4):
                                #This if runs when you input wrong OTP thrice.
                                if k == 3:
                                    print("You made three wrong entries, changing account status to blocked.")
                                    print('Quit Program.')
                                    sh1['G%d'%(i)].value = 'Blocked'
                                    wbook.save("Workbook.xlsx")
                                    flag = 1
                                    break
                                elif flag != 1:
                                    OTP = int(input("Enter OTP: "))
                        
                                    if OTP == x:
                                        print("Please wait you are being redirected.")
                                        
                                        for t in range(4):
                                            
                                            #If user makes any other entries other than 'C' 
                                            #(refers credit) or 'D'(refers debit) thrice, account gets
                                            #blocked.
                                            if t == 3:
                                                print("You made three wrong entries, changing account status to blocked.")
                                                print('Quit Program.')
                                                sh1['G%d'%(i)].value = 'Blocked'
                                                wbook.save("Workbook.xlsx")
                                                flag = 1
                                                break
                                            elif flag != 1:
                                                typeoftrans=input("Enter 'C' for credit and 'D' for Debit: ")
                                                if typeoftrans=='C':
                                                    d1 = {'initial':sh1['F%d'%(i)].value}
                                                    money=int(input("Enter Amount to be credited."))
                                                    if money < 0:
                                                        print("Invalid amount entered to be credited, cannot complete transaction.")
                                                        flag = 1
                                                        break
                                                    sh1['F%d'%(i)].value += money
                                                    localtime = time.asctime( time.localtime(time.time()) )
                                                    s7 = "credit"
                                                    d1[s7] = money
                                                    d1['balance'] = sh1['F%d'%(i)].value
                                                    flag = 1
                                                    wbook.save("Workbook.xlsx")
                                                    break
                                            
                                                elif typeoftrans == 'D':
                                                    d1 = {'initial':sh1['F%d'%(i)].value}
                                                    money = int(input("Enter Amount to be debited: "))
                                                    if money < 0:
                                                        print("Invalid amount entered to be debited, cannot complete transaction.")
                                                        flag = 1
                                                        break
                                                    if sh1['F%d'%(i)].value < money:
                                                        print("Not enough balance, couldn't complete transaction.")
                                                        flag = 1
                                                        break
                                                    sh1['F%d'%(i)].value -= money
                                                    localtime = time.asctime( time.localtime(time.time()) )
                                                    s7 = "debit"
                                                    d1[s7] = money
                                                    d1['balance'] = sh1['F%d'%(i)].value
                                                    flag = 1
                                                    wbook.save("Workbook.xlsx")
                                                    break
                                            
                                                else:
                                                    print("Wrong character entered.\n")
                                            elif flag == 1:
                                                break
                                    
                                    else:
                                        print("Wrong entry.\n")
                                elif flag == 1:
                                        break
                        else:
                            print("Wrong bank account number.\n")
                    elif flag == 1:
                        break
        else:
            if flag != 1:
                print("Sorry no such name exists in our database.\n")     
            else:
                break
    elif flag == 1:
        break

#Prints the mini statement.
if len(d1) > 1:
    print("\n****Mini Statement****")
    print(Name)
    print(localtime)
    for i in d1:
        print(i, d1[i])
