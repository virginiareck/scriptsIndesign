@echo off
chcp 65001

pushd %~dp0

7z.exe e -y -o"%~dP2unzip_%1" %2 >> log.txt

pdftk "%~dP2unzip_%1\*.pdf" cat output "%~dP2%1.pdf" >> log.txt
