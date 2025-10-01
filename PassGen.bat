@echo off
chcp 65001 >nul
setlocal EnableExtensions
title PassGen
color a
cls
echo  ______ _______ _______ _______ _______ _______ _______ 
echo ^|   __ \   _   ^|     __^|     __^|     __^|    ___^|    ^|  ^|
echo ^|    __/       ^|__     ^|__     ^|    ^|  ^|    ___^|       ^|
echo ^|___^|  ^|___^|___^|_______^|_______^|_______^|_______^|__^|____^|
echo.
echo.
echo.

:LENGTH
set "len="
set /p "len=Enter desired length: "
echo %len%| findstr /R "^[0-9][0-9]*$" >nul || (echo Error& goto LENGTH)
if %len% LEQ 0 (echo.& echo Error& goto LENGTH)
if %len% GTR 1024 (echo.& echo Error& goto LENGTH)
goto GENERATION

:GENERATION
set "vbs=%temp%\pvpg_gen.vbs"
set "out=%temp%\pvpg_out.txt"
> "%vbs%"  echo(Option Explicit
>>"%vbs%"  echo(Dim l,upper,lower,digits,special,allChars,arr(),i,j,tmp,out
>>"%vbs%"  echo(If WScript.Arguments.Count^>=1 Then l = CInt(WScript.Arguments(0)) Else l = 16
>>"%vbs%"  echo(upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
>>"%vbs%"  echo(lower = "abcdefghijklmnopqrstuvwxyz"
>>"%vbs%"  echo(digits = "0123456789"
>>"%vbs%"  echo(special = "!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
>>"%vbs%"  echo(Randomize
>>"%vbs%"  echo(ReDim arr(l-1)
>>"%vbs%"  echo(i = 0
>>"%vbs%"  echo(If l ^>= 4 Then
>>"%vbs%"  echo(  arr(i)=Mid(upper, Int(Rnd()*Len(upper))+1, 1) : i=i+1
>>"%vbs%"  echo(  arr(i)=Mid(lower, Int(Rnd()*Len(lower))+1, 1) : i=i+1
>>"%vbs%"  echo(  arr(i)=Mid(digits,Int(Rnd()*Len(digits))+1,1) : i=i+1
>>"%vbs%"  echo(  arr(i)=Mid(special,Int(Rnd()*Len(special))+1,1) : i=i+1
>>"%vbs%"  echo(End If
>>"%vbs%"  echo(allChars = upper ^& lower ^& digits ^& special
>>"%vbs%"  echo(For j = i To l-1
>>"%vbs%"  echo(  arr(j)=Mid(allChars, Int(Rnd()*Len(allChars))+1, 1)
>>"%vbs%"  echo(Next
>>"%vbs%"  echo(For j = l-1 To 1 Step -1
>>"%vbs%"  echo(  i = Int(Rnd()*(j+1))
>>"%vbs%"  echo(  tmp = arr(j) : arr(j) = arr(i) : arr(i) = tmp
>>"%vbs%"  echo(Next
>>"%vbs%"  echo(out = ""
>>"%vbs%"  echo(For i = 0 To l-1 : out = out ^& arr(i) : Next
>>"%vbs%"  echo(WScript.Echo out
cscript //nologo "%vbs%" %len% > "%out%"
echo.
<nul set /p "= Password: "
type "%out%"
type "%out%" | clip
del /q "%vbs%" "%out%" >nul 2>&1
goto AGAIN

:AGAIN
echo.
echo 1) Again
echo 2) Exit
echo.
set "again="
set /p "again= "
if "%again%"=="1" goto LENGTH
if "%again%"=="2" exit
echo.
echo Error
goto AGAIN