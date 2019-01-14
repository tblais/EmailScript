import win32com.client
import pandas as pd

olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")

myData = pd.read_csv("H:\\Projects\\12312018_2019 Recruitment\\Data.csv")

myData.columns = myData.columns.str.replace(' ','_')

for index, row in myData.iterrows():
    print(row['Email'])
    newMail = obj.CreateItem(olMailItem)
    newMail.subject = "2019 Recruitment Planning - " + row['Department']
    newMail.To = row['Email']
    newMail.HTMLBody = """
                <font face="Calibri"><p><strong>Summer Recruitment has begun!</strong></p>
                <p>We are planning to recruit for """ + row['Department'] + """ 
                and want to ensure you're perpared</p><p>We'll be holding meetings. Please attend</p>
                </font>"""
    newMail.display()
