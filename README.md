# banking-platform-code
This project is completely made in python and run in jupyter notebook if it doesn't as .py file by changing the extension 
or by simply copy pasting the code in a single cell of jupyter notebook.
I made use of "openpyxl" library to make changes in excel sheet(this is the sheet where the data is present).
and "smtplib" library for sending emails automatically.


The project is as follows:

Stage 1 : User Validity
Ask user to enter Full Name or User ID
If Full Name / User ID not there then tell invalid User
If it exists, then ask user to enter Bank Account Number
If A/c number doesn't exist, then tell invalid user
If both are correct, then proceed to Stage 2
Give user three tries. If invalid input is given in all three tries then change account status 
from Active to Blocked and print Account Blocked and print quit program in Stage 1.

Stage 2: OTP Generation and Validity
Through code generate a 4 digit random OTP number
Send this OTP to that User via email
Then ask user to enter OTP from mail. Match this OTP with one generated in code.
If OTP matches then go to Stage 3. Else say Invalid OTP.
Give user three tries. If wrong OTP in all three tries then change account status
from Active to Blocked and print Account Blocked and print quit program in Stage 2

Stage 3 : Choice of Transaction
Ask user which type of transaction he wishes to do. ‘C’ for credit money and ‘D’ for debit money
If User enters any other input here then handle error and print Invalid Input.
Again give user three tries and if all wrong then change account status from
Active to Blocked and print Account Blocked and print quit program in Stage 3
If correct input then print current balance
Then ask user for amount to be credited or debited
Add or subtract amount entered by user to the current account balance
and print the final account balance and write the same to account balance cells in excel workbook.
Print “Successful Transaction of Rs.__ completed”

There will be five cases here:
Case 1-> When you input name wrong three times
Case 2-> When you input account number wrong three times.
Case 3-> When you input OTP number wrong three times.
Case 4-> When you input wrong entries for credit or debit three times.
Case 5-> When you run correctly.

Screenshots for each case:
Case 1:
![Case 1](https://github.com/tusshar2000/banking-platform-code/blob/master/screenshots/case1.PNG)

Case 2:
![Case 2](https://github.com/tusshar2000/banking-platform-code/blob/master/screenshots/case2.PNG)

Case 3:
![Case 3](https://github.com/tusshar2000/banking-platform-code/blob/master/screenshots/case3.PNG)

Case 4:
![Case 4](https://github.com/tusshar2000/banking-platform-code/blob/master/screenshots/case4.PNG)

Case 5:
![Case 5](https://github.com/tusshar2000/banking-platform-code/blob/master/screenshots/case5.PNG)

Excel sheet at the starting.

Excel sheet at the ending.

