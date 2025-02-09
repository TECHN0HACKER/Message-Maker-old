@echo off
title MESSAGE MAKER
:1
cls
color 1
echo Welcome!
pause
cls
echo Type the serial number of the kind of message box you want 
echo 1) Normal
echo 2) Error
echo 3) Information
echo 4) Question
echo 5) Warning
echo 6) Icon on the title bar
echo 7) Message (Only for Windows 10 Pro and above)
echo 8) Message
echo 9) Input Box
set /p option=
if %option%==1 goto Normal
if %option%==2 goto Error
if %option%==3 goto Info
if %option%==4 goto Question
if %option%==5 goto Warning
if %option%==6 goto Icon
if %option%==7 goto Msg
if %option%==8 goto Message
if %option%==9 goto InputBox
pause

:Normal
cls
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
color 2
echo Type the serial number of the buttons to appear on the box
echo 1) OK
echo 2) OK and Cancel
echo 3) Abort, Retry and Ignore
echo 4) Yes, No and Cancel
echo 5) Yes and No
echo 6) Retry and Cancel
set /p check=
if %check%==1 goto ok
if %check%==2 goto oc
if %check%==3 goto ari
if %check%==4 goto ync
if %check%==5 goto yn
if %check%==6 goto rc
cls

:ok
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",0,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:oc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",1,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ari
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",2,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ync
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",3,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:yn
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",4,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:rc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",5,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:Error
cls
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
color 2
echo Type the serial number of the buttons to appear on the box
echo 1) OK
echo 2) OK and Cancel
echo 3) Abort, Retry and Ignore
echo 4) Yes, No and Cancel
echo 5) Yes and No
echo 6) Retry and Cancel
set /p check=
if %check%==1 goto ok
if %check%==2 goto oc
if %check%==3 goto ari
if %check%==4 goto ync
if %check%==5 goto yn
if %check%==6 goto rc
cls

:ok
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",16,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:oc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",17,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ari
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",18,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ync
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",19,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:yn
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",20,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:rc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",21,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:Info
cls
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
color 2
echo Type the serial number of the buttons to appear on the box
echo 1) OK
echo 2) OK and Cancel
echo 3) Abort, Retry and Ignore
echo 4) Yes, No and Cancel
echo 5) Yes and No
echo 6) Retry and Cancel
set /p check=
if %check%==1 goto ok
if %check%==2 goto oc
if %check%==3 goto ari
if %check%==4 goto ync
if %check%==5 goto yn
if %check%==6 goto rc
cls

:ok
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",64,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:oc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",65,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ari
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",66,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ync
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",67,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:yn
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",68,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:rc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",69,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:Question
cls
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
color 2
echo Type the serial number of the buttons to appear on the box
echo 1) OK
echo 2) OK and Cancel
echo 3) Abort, Retry and Ignore
echo 4) Yes, No and Cancel
echo 5) Yes and No
echo 6) Retry and Cancel
set /p check=
if %check%==1 goto ok
if %check%==2 goto oc
if %check%==3 goto ari
if %check%==4 goto ync
if %check%==5 goto yn
if %check%==6 goto rc
cls

:ok
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",32,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:oc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",33,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ari
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",34,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ync
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",35,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:yn
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",36,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:rc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",37,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:Warning
cls
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
color 2
echo Type the serial number of the buttons to appear on the box
echo 1) OK
echo 2) OK and Cancel
echo 3) Abort, Retry and Ignore
echo 4) Yes, No and Cancel
echo 5) Yes and No
echo 6) Retry and Cancel
set /p check=
if %check%==1 goto ok
if %check%==2 goto oc
if %check%==3 goto ari
if %check%==4 goto ync
if %check%==5 goto yn
if %check%==6 goto rc
cls

:ok
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",48,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:oc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",49,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ari
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",50,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ync
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",51,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:yn
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",52,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:rc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",53,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:Msg
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
cls
color 2
set /p message=Type the message here
cls
set /p speak=Do you want your message to be spoken? (y/n)
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
cls
echo Done! your message box is ready now!
pause
cls
msg * %message%
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:Message
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
cls
color 2
set /p message=Type the message here
cls
set /p speak=Do you want your message to be spoken? (y/n)
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
cls
echo WScript.echo "%message%" >>message.vbs
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:InputBox
cls
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
color 2
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=InputBox("%message%","%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:Icon
cls
if exist "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs" (
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
) else (
cd "%LOCALAPPDATA%"
)
if exist "%LOCALAPPDATA%\MESSAGE MAKER" (
cd "%LOCALAPPDATA%\MESSAGE MAKER"
) else (
mkdir "%LOCALAPPDATA%\MESSAGE MAKER"
cd "%LOCALAPPDATA%\MESSAGE MAKER"
)
color 2
echo Type the serial number of the buttons to appear on the box
echo 1) OK
echo 2) OK and Cancel
echo 3) Abort, Retry and Ignore
echo 4) Yes, No and Cancel
echo 5) Yes and No
echo 6) Retry and Cancel
set /p check=
if %check%==1 goto ok
if %check%==2 goto oc
if %check%==3 goto ari
if %check%==4 goto ync
if %check%==5 goto yn
if %check%==6 goto rc
cls

:ok
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",4096,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:oc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",4097,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ari
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",4098,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:ync
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",4099,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:yn
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",4100,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit

:rc
cls
set /p message=Type the message here
cls
set /p title=Type the title of the message box
cls
echo Do you want your message to be spoken? (y/n)
set /p speak=
if %speak%==y goto next
if not %speak%==y goto back
cls
:next
cls
echo Type the serial number of the speaker:
echo 1) Microsoft David
echo 2) Microsoft Zira
set /p speaker=
if %speaker%==1 goto David
if %speaker%==2 goto Zira

:David
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(0) >>message.vbs
goto speedcheck

:Zira
echo Dim spV >message.vbs
echo Set spV = CreateObject("SAPI.spVoice") >>message.vbs
echo Set spV.Voice = spV.GetVoices.Item(1) >>message.vbs
goto speedcheck

:speedcheck
cls
echo Type the speaking speed. The numbers can be below zero (eg: -1, -2, -3... etc)
set /p speed=
echo spV.Rate=%speed% >>message.vbs
echo spV.Speak "%title%,%message%" >>message.vbs
:back
echo x=msgbox("%message%",4101,"%title%") >>message.vbs
cls
echo Done! your message box is ready now!
pause
cls
start message.vbs
echo Exit? (y/n)
set /p exit=
if %exit%==n goto no
if not %exit%==n goto yes
pause

:no
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
goto 1

:yes
del "%LOCALAPPDATA%\MESSAGE MAKER\message.vbs"
exit
