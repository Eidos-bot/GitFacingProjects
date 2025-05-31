import win32com.client as client
from Auto_Email_Attach import converting_excel, next_monday


upcoming_monday = next_monday("Today",6,weeknum=0, free_date="No")

upcoming_due = next_monday(upcoming_monday.replace("Monday, ", ""),11,weeknum=0, free_date="Yes")

last_subday = next_monday("Today",0,weeknum=1,free_date="No")

outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.Display()
message.To = input('user_email')
message.CC = ""
message.Subject = "AP Alert Test"
message.Attachments.Add(Source=converting_excel())
messagetest = "Hi guys this email was done through python. See I have my signature? So it should look like a normal email now."
messagebody = f"""

    Good morning, <br>
    <br>
    AP is now processing invoices that will be paid on our {upcoming_monday} <span style=color:red'>ACH </span> payment run with a due date up to: {upcoming_due}. Please submit any invoices you wish to be paid in this run by EOD {last_subday}. <br>
    <br>
    <span style=color:red'> Check payments are disbursed the 2nd Tuesday and last Tuesday of the month. </span> <br>
    <br>
    Thank you, <br>
    
"""



message.HTMLbody = f"""<!DOCTYPE html>

<p>{messagebody}</p>

<p>
    <span> <b style='font-family:"Times New Roman",serif;color:maroon'>Christopher Dessources </b> </span> <br>

    <span style='font-size:10.0pt;font-family:"Times New Roman",serif'>Accounts Payable Specialist </span> <br>

    <span> <b style='font-size:10.0pt;font-family:"Times New Roman",serif;color:maroon'>Brooklyn Law School </b> </span> <br>

    <span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:maroon'>Phone: </span> <span style='font-size:10.0pt;font-family:"Times New Roman",serif'>718-780-0340 </span> <br>

    <span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:maroon'>Email: </span> <span style='font-size:10.0pt;font-family:"Times New Roman",serif'><a href="mailto:christopher.dessourc@brooklaw.edu">christopher.dessourc@brooklaw.edu</a></span> <br>
</p>"""

message.Send()


