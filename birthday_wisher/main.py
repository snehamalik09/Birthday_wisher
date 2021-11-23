#Author - Sneha Malik
#Date - Nov 2021

import pandas 
import datetime
import smtplib

#add your own authenticated credentials here
gmail_id=''
gmail_password=''


def send_email(to, sub, msg):
    s=smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(gmail_id, gmail_password)
    s.sendmail(gmail_id, to, f"SUBJECT - {sub}\n\n{msg}")
    s.quit()       
    print(f"Email sent to {to} with subject {sub} and msg {msg}")
    
   
if __name__ == '__main__':
    #reading the excel sheet
    df=pandas.read_excel("data.xlsx")
    today=datetime.datetime.now().strftime("%d-%m")
    wished_year=dict()
    writeInd=[]
    
    #iterating over rows of excel
    for index, item in df.iterrows():
        #changing birthdate in day-month format
        bdate=item["Birthdate"].strftime("%d-%m")    
        present_year=datetime.datetime.now().strftime("%Y")
        
        if today==bdate and present_year not in str(item["Year"]):
            send_email(item['Email id'], 'Happy birthday', item['Dialogue'])
            writeInd.append(index)
    
#appending "Year" column in excel sheet with present year on which we have wished the birthday to avoid sending mail again and again.        
for i in writeInd:
    yr=df.loc[i, "Year"]
    df.loc[i, "Year"]=str(yr) + "," + str(present_year)
    
df.to_excel("data.xlsx", index=False)


            
            
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
            
#The strftime() function is used to convert date and time objects to their string representation. It takes one or 
# more input of formatted code and returns the string representation.
#for reading we did pip install xlrd
#for writing we will do pip install xlwt
#to_excel() uses a library called xlwt and openpyxl internally.

#xlwt is used to write .xls files (formats up to Excel2003)
#openpyxl is used to write .xlsx (Excel2007 or later formats).
#by first converting it into a Pandas DataFrame and then writing the DataFrame to Excel.
#loc attribute access a group of rows and columns by label(s) or a boolean array in the given DataFrame. 
#Example #1: Use DataFrame. loc attribute to access a particular cell in the given Dataframe using the index
# and column labels.