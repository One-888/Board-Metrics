@ECHO OFF
cls
ECHO ===========================================
ECHO Welcome to PGH2O Access for PowerBI
ECHO Copyright 2024 PGH2O, All rights reserved.
ECHO ===========================================
TIMEOUT /T 1 /nobreak > NUL
ECHO Copying file..
TIMEOUT /T 1 /nobreak > NUL
ECHO Starting.. Access
TIMEOUT /T 1 /nobreak > NUL
ECHO .
ECHO This windows will exit automatically after Access Database exits.
C:
CD %USERPROFILE%\Downloads
robocopy "\\fs1\Shared\PWSA\Access_for_PowerBI" "%USERPROFILE%\Downloads" Admin.accdb /njs /njh
%USERPROFILE%\Downloads\Admin.accdb

REM pause