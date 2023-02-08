
#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
comm:=","
quot="
thing='

whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")


if FileExist("c:\users\public\documents\archive submission\software\checksum.ini")
{}
else
{
whr.Open("GET", "https://raw.githubusercontent.com/BOSTLI/ArchiveSubmissionSystem/main/checksum.ini", true)
whr.Send()
; Using 'true' above and the call below allows the script to remain responsive.
whr.WaitForResponse() ;this is taken from the installer. Can also be located as an example on the urldownloadtofile page of the quick reference guide.
version := whr.ResponseText
fileappend,%version%,c:\users\public\documents\archive submission\software\checksum.ini
}

if FileExist("c:\users\public\documents\archive submission\software\csjarchives3.png")
{}
else
{
url:="https://github.com/BOSTLI/ArchiveSubmissionSystem/raw/main/csjarchives3.png"
Filename = checksum verify v1.1
UrlDownloadToFile, *0 %url%,c:\users\public\documents\archive submission\software\csjarchives3.png
}

if FileExist("c:\users\public\documents\archive submission\software\checksum verify v1.1.exe")
{}
else
{
url:="https://github.com/BOSTLI/ArchiveSubmissionSystem/blob/main/checksum%20verify%20v1.1.exe?raw=true"
Filename = checksum verify v1.1
UrlDownloadToFile, *0 %url%,c:\users\public\documents\archive submission\software\checksum verify v1.1.exe
}

if FileExist("c:\users\public\documents\archive submission\software\checksum verify v1.2.exe")
{}
else
{
url:="https://github.com/BOSTLI/ArchiveSubmissionSystem/blob/main/checksum%20verify%20v1.2.exe?raw=true"
Filename = checksum verify v1.2
UrlDownloadToFile, *0 %url%,c:\users\public\documents\archive submission\software\checksum verify v1.2.exe
}

if FileExist("c:\users\public\documents\archive submission\software\checksum.exe")
{}
else
{
url:="https://github.com/BOSTLI/ArchiveSubmissionSystem/blob/main/checksum.exe?raw=true"
Filename = checksum verify v1.1
UrlDownloadToFile, *0 %url%,c:\users\public\documents\archive submission\software\checksum.exe
}


if FileExist("c:\users\public\documents\archive submission\software\swithmail.exe")
{}
else
{
url:="https://github.com/BOSTLI/ArchiveSubmissionSystem/blob/main/SwithMail.exe?raw=true"
Filename = swithmail.exe
UrlDownloadToFile, *0 %url%,c:\users\public\documents\archive submission\software\swithmail.exe
}

if FileExist("c:\users\public\documents\archive submission\software\email-data.txt")
{}
else
{
url:="https://raw.githubusercontent.com/BOSTLI/ArchiveSubmissionSystem/main/Email-Data.txt"
Filename = email-data.txt
UrlDownloadToFile, *0 %url%, c:\users\public\documents\archive submission\software\email-data.txt
}

if FileExist("c:\users\public\documents\archive submission\software\log.txt")
FileDelete, c:\users\public\documents\archive submission\software\log.txt
if FileExist("c:\users\public\documents\archive submission\software\processed.txt")
FileDelete, c:\users\public\documents\archive submission\software\processed.txt

if FileExist("c:\users\public\documents\archive submission\software\userinformation.txt")
{
FileRead, var, c:\users\public\documents\archive submission\software\userinformation.txt
n1=1
Loop, parse, var, %comm%
{
if n1=1
depositer:=A_LoopField
else if n1=2
email:=A_LoopField
n1:=n1+1
}
}
else
{
InputBox, email,Email Address, Please fill in your email address below.
InputBox, depositer,Depositer Info, Please fill in your name below.
FileAppend,%depositer%%comm%%email%,c:\users\public\documents\archive submission\software\userinformation.txt
}

Loop
{
if email=
{
InputBox, email,Email Address, Please fill in your email address below.
InputBox, depositer,Depositer Info, Please fill in your name below.
FileAppend,%depositer%%comm%%email%,c:\users\public\documents\archive submission\software\userinformation.txt
break
}
Else
break
}

FileRead, script, script.txt
script:= StrReplace(script,"blalande","Archive")
script:= StrReplace(script,"put " quot "c:\users\public\documents\archive submission\submission\script.txt" quot,"")
script:= StrReplace(script,"cd /Injest","get processed.txt")
;script:= StrReplace(script,"close", depositer datet ".tar")
if FileExist("c:\users\public\documents\archive submission\software\script2.txt")
FileDelete, c:\users\public\documents\archive submission\software\script2.txt
FileAppend,%script%,script2.txt
run, winscp.com /script=script2.txt

loop
{
Process, Exist, winscp.com
If ErrorLevel = 0
{
break
}
Else
{
;MsgBox, Notepad is running.
}
}

Loop, read, processed.txt
{
if InStr(A_LoopReadLine,depositer)    
{
if FileExist("c:\users\public\documents\archive submission\submission\" A_LoopReadLine)
FileRemoveDir, c:\users\public\documents\archive submission\submission\%A_LoopReadLine%, 1
if FileExist("c:\users\public\documents\archive submission\submission\" A_LoopReadLine ".tar")
FileDelete, c:\users\public\documents\archive submission\submission\%A_LoopReadLine%.tar
}
}



Gui +hWndhMainWnd
Gui Add, Picture, x-4 y-340 w985 h468, c:\users\public\documents\archive submission\software\csjarchives3.png
Gui Add, Picture, x-17 y326 w0 h0, Who is transferring the files? (Depositer)
Gui Font, s12 Bold c0x3591DB
Gui Add, Text, x60 y153 w310 h21 +0x200 vtext1, Who made the original files? *
Gui Font
Gui Font, s12
Gui Add, Edit, x60 y187 w313 h38
Gui Font, s12 Bold c0x3591DB
Gui Add, Text, x60 y241 w400 h41 vtext2, Who is submitting the files?*
Gui Font
Gui Font, s12
Gui Add, Edit, x60 y290 w314 h40
Gui Font
Gui Font, s12 Bold c0x3591DB
Gui Add, Text, x544 y156 w420 h21 +0x200 vtext3, What department do the files belong too? *
Gui Font
Gui Font, s12
Gui Add, DropDownList,  x541 y188 w333 h350 vdepo,Leadership|Administration|Finance|Human Resources|Nursing Care|Associates/Companions|Casa Maria Refugee Homes|CSJ Spirituality Centre|Spiritual Ministries Network|Stillpoint House of Prayer|St. Josephs Hospitality Centre|Individual Sister|Vocation|Local Residence|Archives Donation|Committee|Upper Room House of Prayer|Villa St. Joseph
;Gui Add, Edit, x541 y388 w333 h38 vdepo
Gui Font, s12 Bold
Gui Add, Text, x541 y427 w413 h21 +0x200 vtext4, Are there any access restrictions?
Gui Font
Gui Font, s12
;Gui Add, Edit, x541 y472 w335 h116
Gui Font, s12 Bold
Gui Add, Text, +Wrap x60 y340 w363 h49 +0xC vtext5, Are there any copyright protected files? E.g.%comm% if photos who created the photos? 
Gui Font, s12 Bold c0x3591DB
Gui Add, Text, +Wrap x60 y340 w363 h49 +0xC vtext52, Are there any copyright protected files? E.g.%comm% if photos who created the photos? *
Gui Add, DropDownList,  x60 y385 w333 h350 Choose1 vcopyr gcopyri,No|Yes
Gui Font
Gui Font, s12
Gui Add, Edit, x60 y415 w335 h80 v1t
Gui, Add, Progress, x60 y415 w335 h80 cDEDDDD vMyProgress2, 100
Gui, Add, Progress, x541 y300 w385 h115 cDEDDDD vMyProgress3, 100
Gui Font, s12 Bold
Gui Add, Text, x541 y238 w413 h21 +0x200 vtext6, Is there any sensitive information in the files?

Gui Font, s12 Bold c0x3591DB
Gui Add, Text, x541 y238 w413 h21 +0x200 vtext62, Is there any sensitive information in the files? *
Gui Font, s12 Bold
Gui Add, Text, x60 y500 w783 h21 +0x200 vtext7, Please briefly describe what files you are submitting | e.g. %comm% meeting minutes%comm% newletters%comm% etc. *
Gui Font
Gui Font, s12
Gui Add, DropDownList,  x541 y270 w333 h350 Choose1 vsensinfo gsensin,No|Yes
Gui Add, CheckBox, x541 y300 w150 h53 vbutt1, Health information
Gui Add, CheckBox, x541 y360 w150 h53 vbutt2, Social insurance information
Gui Add, CheckBox, x701 y300 w250 h53 vbutt3, Financial information including bank account information
Gui Add, CheckBox, x701 y360 w250 h53 vbutt4, Contact information (emails, phone numbers, addresses)


Gui Add, DropDownList,  x541 y452 w333 h350 Choose1 vaccess,No|Yes
;Gui Add, Edit, x540 y652 w335 h59
Gui Font, s16
Gui Add, Button, x856 y620 w115 h46 gsubmit, Submit
Gui Add, Button, x5 y620 w115 h46 gquit, Quit


Gui Font
Gui Font, s12
Gui Add, Edit, x60 y530 w784 h80 v2t

Gui Font, s40 Bold c0x3591DB
Gui Add, Text, x1 y203 w980 h259 +0x1 vloading, Files are Uploading. Please Wait.
GuiControl, hide, loading
GuiControl, hide, 1t
GuiControl, hide, butt1
GuiControl, hide, butt2
GuiControl, hide, butt3
GuiControl, hide, butt4
GuiControl, hide, text62
GuiControl, hide, text52
Guicontrol, text, 1t, Copyright Holder:



;GuiControl,, depo, %depositer%
Gui Show, w980 h674, CSJ Archives Submission System.

Return

quit:
GuiEscape:
GuiClose:
    ExitApp

sensin:
guicontrolget, yn,, sensinfo
if yn=Yes
{
GuiControl, show, butt1
GuiControl, show, butt2
GuiControl, show, butt3
GuiControl, show, butt4
GuiControl, show, text62
GuiControl, hide, MyProgress3
}    
else
{
GuiControl, hide, butt1
GuiControl, hide, butt2
GuiControl, hide, butt3
GuiControl, hide, butt4
GuiControl, show, MyProgress3
GuiControl, hide, text62
}
return

copyri:
guicontrolget, yn,, copyr
if yn=Yes
{
GuiControl, show, 1t
GuiControl, hide, MyProgress2
GuiControl, show, text52
}    
else
{
GuiControl, hide, 1t
GuiControl, show, MyProgress2
GuiControl, hide, text52
}
return

submit:
guicontrolget, creator,, Edit1
guicontrolget, custody,, Edit2
guicontrolget, description,, 2t
guicontrolget, category,, depo

guicontrolget, copyright,, 1t
guicontrolget, access,, access

GuiControlGet, check,, butt1
if check=1
ISS:="Health information "
GuiControlGet, check,, butt2
if check=1
ISS:=Iss "Social insurance information "
GuiControlGet, check,, butt3
if check=1
ISS:=Iss "Financial information "
GuiControlGet, check,, butt4
if check=1
ISS:=Iss "Contact information "

;msgbox %iss%

guicontrolget, yn,, sensinfo
if yn=Yes
{
if iss=
{
msgbox, 266256,, Missing Sensitive Information, Please check one of the checkboxes.
return
}
}

guicontrolget, yn,, copyr
if yn=Yes
{
if copyright=Copyright Holder:
{
msgbox, 266256,, Missing Copyright information, please fill in all * required fields.
return
}
}

if creator=
{
msgbox, 266256,, Missing creator information, please fill in all * required fields.
return
}
if custody=
{
msgbox, 266256,, Missing custody information, please fill in all * required fields.
return
}
if category=
{
msgbox, 266256,, Missing category information, please fill in all * required fields.
return
}
if description=
{
msgbox, 266256,, Missing description information, please fill in all * required fields.
return
}


GuiControl, hide, edit1
GuiControl, hide, edit2
GuiControl, hide, edit3
GuiControl, hide, edit4
GuiControl, hide, edit5
GuiControl, hide, edit6
GuiControl, hide, text1
GuiControl, hide, text2
GuiControl, hide, text3
GuiControl, hide, text4
GuiControl, hide, text5
GuiControl, hide, text52
GuiControl, hide, text6
GuiControl, hide, text7
GuiControl, hide, text62
guicontrol, show, loading
GuiControl, hide, butt1
GuiControl, hide, butt2
GuiControl, hide, butt3
GuiControl, hide, butt4
GuiControl, hide, butt5
GuiControl, hide, depo
GuiControl, hide, Submit
GuiControl, hide, Quit
GuiControl, hide, 1t
GuiControl, hide, myprogress2
GuiControl, hide, myprogress3
GuiControl, hide, sensinfo
GuiControl, hide, copyr
GuiControl, hide, access


FormatTime, datet,,yyyymmDDhhmiss
FileCreateDir, c:\users\public\documents\archive submission\submission\%depositer%%datet%
FileCreateDir, c:\users\public\documents\archive submission\submission\%depositer%%datet%\metadata
FileMoveDir, c:\users\public\documents\archive submission\Archive Transfer, c:\users\public\documents\archive submission\submission\%depositer%%datet%\Files

Fileappend,Depositer%comm%Email%comm%Creator%comm%Category%comm%Custody%comm%PII%comm%Copyright%comm%Access%comm%Description `n,c:\users\public\documents\Archive Submission\Submission\%depositer%%datet%\metadata\Submission Information.csv
Fileappend,%Depositer%%comm%%Email%%comm%%Creator%%comm%%category%%comm%%Custody%%comm%%PII%%comm%%Copyright%%comm%%Access%%comm%%description%,c:\users\public\documents\Archive Submission\Submission\%depositer%%datet%\metadata\Submission Information.csv

runwait, dirhash.exe %quot%c:\users\public\documents\archive submission\Submission\%depositer%%datet%\files%quot% -t %quot%c:\users\public\documents\archive submission\Submission\%depositer%%datet%\metadata\checksums.txt%quot%
runwait, checksum.exe cqry2t %quot%c:\users\public\documents\archive submission\Submission\%depositer%%datet%\files%quot%
;runwait, 7za.exe a %quot%c:\users\public\documents\archive submission\submission\%depositer%%datet%.zip%quot% %quot%c:\users\public\documents\archive submission\Archive Transfer\%quot%
runwait, cmd.exe /c %quot%%quot%c:\users\public\documents\archive submission\software\Filelist.exe%quot% /USECOLUMNS NAME%comm%EXTENSION%comm%SIZE%comm%FULLPATH%comm%DIRLEVEL%comm%LASTACCESS%comm%LASTCHANGE%comm%CREATIONDATE%comm%MD5%comm%SHA256 %quot%C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\Files\%quot% > %quot%C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\metadata\File Manifest.CSV%quot%%quot%
filecopy, c:\users\public\documents\archive submission\software\checksum verify v1.2.exe, C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\checksum verify v1.2.exe
;filecopy, c:\users\public\documents\archive submission\software\dirhash.exe, C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\dirhash.exe

runwait, 7za.exe a %quot%c:\users\public\documents\archive submission\submission\%depositer%%datet%.tar%quot%  %quot%c:\users\public\documents\archive submission\Submission\%depositer%%datet%\*%quot%

FileRead, script, script.txt
script:= StrReplace(script,"blalande","Archive")
script:= StrReplace(script,"script.txt", depositer datet ".tar")
script:= StrReplace(script,"Injest", "Ingest")
;script:= StrReplace(script,"close", depositer datet ".tar")
if FileExist("c:\users\public\documents\archive submission\software\script2.txt")
FileDelete, c:\users\public\documents\archive submission\software\script2.txt
FileAppend,%script%,script2.txt
run, winscp.com /script=script2.txt /log=log.txt


loop
{
Process, Exist, winscp.com
If ErrorLevel = 0
{
break
}
Else
{
;MsgBox, Notepad is running.
}
}
sleep, 500

FileRead, log, log.txt
if instr(log,"Upload successful")
{


FileRead, emaildata, email-data.txt
emaildata:= StrReplace(emaildata,"{submission-number}", depositer datet)
emaildata:= StrReplace(emaildata,"{sender email address}",email)
emaildata:= StrReplace(emaildata,"{passchange}","8EhyYNOCQ3Pbzc8J0GzhBw")
emaildata:= StrReplace(emaildata,"{attachment}","c:\users\public\documents\archive submission\submission\"  depositer datet "\metadata\file manifest.csv")
if FileExist("c:\users\public\documents\archive submission\software\SwithMailSettings.xml")
FileDelete, c:\users\public\documents\archive submission\software\SwithMailSettings.xml
FileAppend,%emaildata%,SwithMailSettings.xml
run, "c:\users\public\documents\archive submission\software\SwithMail.exe" /s /x "c:\users\public\documents\archive submission\software\SwithMailSettings.xml"



FileCreateDir, c:\users\public\documents\archive submission\Archive Transfer
;run, C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\File Manifest.CSV
msgbox, 266208,, Files have been successfully submitted. You will receive an email with your file manifest shortly.

}
else
failure=1

if failure=1
{ 
FileMoveDir, c:\users\public\documents\archive submission\submission\%depositer%%datet%\Files, c:\users\public\documents\archive submission\Archive Transfer
msgbox, 266256,, Files have NOT been uploaded. Please re-run the application to properly submit your files.
}

exitapp

