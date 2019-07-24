ECHO ON
REM A batch script to execute a Python script
Title Check Transaction
REM address of python folder in local computer
SET PYTHONFOLDER=C:\Users\melody\AppData\Local\Programs\Python\Python36-32
SET PATH=%PATH%;%PYTHONFOLDER%
SET FILEFOLDER=%PYTHONFOLDER%\mytools\transactionCheck\
python %FILEFOLDER%main.py

PAUSE