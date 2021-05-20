import win32com.client as client

outlook = client.Dispatch("Outlook.Application")

CemaEmail = open('Toscano Files/CemaEmail.html', 'r').read()
CoopEmail = open('Toscano Files/CoopEmail.html', 'r').read()

message = outlook.CreateItem(0)
message.To = 'chrisheyer0@gmail.com'
message.CC = 'chrisheyer0@gmail.com; cheyer@tatpc.com'
message.Subject = 'derp'
message.HTMLBody = CemaEmail
message.Send()
