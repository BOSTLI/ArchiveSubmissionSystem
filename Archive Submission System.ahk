
#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
comm:=","
quot="

msgbox Hello World

Gui +hWndhMainWnd +AlwaysOnTop -Caption +Owner
Gui Add, Picture, x-4 y-320 w985 h651, csjarchives2.png
Gui Add, Picture, x-17 y326 w0 h0, Who is transferring the files? (Depositer)
Gui Font, s12
Gui Add, Text, x80 y353 w310 h21 +0x200 vtext1, Who created the files? (CREATOR) *
Gui Font
Gui Font, s12
Gui Add, Edit, x77 y387 w313 h38
Gui Font, s12
Gui Add, Text, x75 y441 w450 h21 +0x200 vtext2, How did you receive the files from the creator? (CUSTODY)
Gui Font
Gui Font, s12
Gui Add, Edit, x75 y478 w314 h40
Gui Font
Gui Font, s12
Gui Add, Text, x544 y356 w320 h21 +0x200 vtext3, Who is submitting the files? (DEPOSITOR) *
Gui Font
Gui Font, s12
Gui Add, Edit, x541 y388 w333 h38 vdepo
Gui Font, s12
Gui Add, Text, x541 y607 w413 h21 +0x200 vtext4, Are there any access restrictions? (ACCESS)
Gui Font
Gui Font, s12
Gui Add, Edit, x541 y472 w335 h116
Gui Font, s12
Gui Add, Text, +Wrap x72 y522 w363 h49 +0xC vtext5, Are there any copyright protected files? E.g. if photos – who created the photos? 
Gui Font
Gui Font, s12
Gui Add, Edit, x71 y583 w335 h108
Gui Font, s12
Gui Add, Text, x541 y438 w413 h21 +0x200 vtext6, Is there any sensitive information in the files? (PII)
Gui Font
Gui Font, s12
Gui Add, Edit, x540 y632 w335 h59
Gui Font, s16
Gui Add, Button, x856 y720 w115 h46 gsubmit, Submit
Gui Add, Button, x5 y720 w115 h46 gquit, Quit

Gui Font, s40
Gui Add, Text, x1 y403 w980 h259 +0x200 vloading, Files are Uploading. Please Wait.
GuiControl, hide, loading

Gui Show, w980 h774, CSJ Archives Submission System.

if FileExist ("archiveconfig.txt")
{
FileRead, depositer,  archiveconfig.txt
GuiControl,, depo, %depositer%
}

if FileExist ("archiveconfigemail.txt")
{
FileRead, email,  archiveconfigemail.txt
}

Return

quit:
GuiEscape:
GuiClose:
    ExitApp

submit:
guicontrolget, creator,, Edit1
guicontrolget, custody,, Edit2
guicontrolget, depositer,, Edit3
guicontrolget, PII,, Edit4
guicontrolget, copyright,, Edit5
guicontrolget, access,, Edit6  


if creator=
{
msgbox, 266256,, Missing creator information, please fill in all * required fields.
return
}
;if custody=
;{
;msgbox, 266256,, Missing custody information, please fill in all * required fields.
;return
;}
if depositer=
{
msgbox, 266256,, Missing depositer information, please fill in all * required fields.
return
}

if FileExist ("archiveconfig.txt")
{}
else
{
FileAppend,%depositer%, archiveconfig.txt
}

FormatTime, datet,,yyyymmDDhhmiss
FileCreateDir, c:\users\public\documents\archive submission\submission\%depositer%%datet%
FileMoveDir, c:\users\public\documents\archive submission\Archive Transfer, c:\users\public\documents\archive submission\submission\%depositer%%datet%\Files

Fileappend,Depositer%comm%Email%comm%Creator%comm%Custody%comm%PII%comm%Copyright%comm%Access `n,c:\users\public\documents\Archive Submission\Submission\%depositer%%datet%\Submission Information.csv
Fileappend,%Depositer%%comm%%Email%%comm%%Creator%%comm%%Custody%%comm%%PII%%comm%%Copyright%%comm%%Access%,c:\users\public\documents\Archive Submission\Submission\%depositer%%datet%\Submission Information.csv

runwait, dirhash.exe %quot%c:\users\public\documents\archive submission\Submission\%depositer%%datet%\files%quot% -t %quot%c:\users\public\documents\archive submission\Submission\%depositer%%datet%\checksums.txt%quot%
;runwait, 7za.exe a %quot%c:\users\public\documents\archive submission\submission\%depositer%%datet%.zip%quot% %quot%c:\users\public\documents\archive submission\Archive Transfer\%quot%
runwait, cmd.exe /c c:\archive\Filelist.exe /USECOLUMNS NAME%comm%EXTENSION%comm%SIZE%comm%FULLPATH%comm%DIRLEVEL%comm%LASTACCESS%comm%LASTCHANGE%comm%CREATIONDATE%comm%MD5%comm%SHA256 %quot%C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\Files\%quot% > %quot%C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\File Manifest.CSV%quot%
filecopy, c:\archive\checksum verify.exe, C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\Checksum Verify.exe
filecopy, c:\archive\dirhash.exe, C:\Users\Public\Documents\Archive Submission\Submission\%depositer%%datet%\dirhash.exe
FileRead, script, script.txt
script:= StrReplace(script,"script.txt", depositer datet)
if FileExist("script2.txt")
FileDelete, script2.txt
FileAppend,%script%,script2.txt
run, winscp.com /script=script2.txt

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
GuiControl, hide, text6
guicontrol, show, loading

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
FileMove, c:\users\public\documents\archive submission\submission\%depositer%%datet%.zip, c:\users\public\documents\archive submission\backups\%depositer%%datet%.zip
runwait, cmd.exe /c rmdir /s /q %quot%c:\users\public\documents\archive submission\Archive Transfer%quot%
sleep,200
FileCreateDir, c:\users\public\documents\archive submission\Archive Transfer
msgbox, 266208,, Files have been successfully submitted.

exitapp

