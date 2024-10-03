@echo off
cd /d %~dp0
echo �������J�n���܂�

rem chocolatey���Ȃ����Wget�Ƃ��킹�ăC���X�g�[������
where /q choco
if %errorlevel% == 0 ( :: �R�}���h�����݂����
    echo choco exists
) else (               :: �R�}���h�����݂��Ȃ����
     Y |"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
     Y |choco install Wget
)

rem Chrome�̃o�[�W�������m�F���āA�����o�[�W������ChromeDriver���_�E�����[�h
echo webdriver download

rem set version=dir /B /AD "C:\Program Files\Google\Chrome\Application"
set versionPath="C:\Program Files\Google\Chrome\Application"
for /f "usebackq" %%A in (`dir /B /AD %versionPath%`) do set version=%%A&goto :exit_for
:exit_for
echo version is %version%
set url=https://storage.googleapis.com/chrome-for-testing-public/%version%/win64/chromedriver-win64.zip
wget %url%

rem �_�E�����[�h����zip�t�@�C������
echo unzip download file
set zipFilePath=.\chromedriver-win64.zip
set destFolderPath=.\output

set psCommand=powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Path %zipFilePath% -DestinationPath %destFolderPath% -Force
%psCommand%

rem C:\ProgramData�ȉ���scrapingFanbox�t�H���_�����AChromeDriver���R�s�[
echo webdriver move stanby

set outPutPath=C:\ProgramData\scrapingFanbox
if not exist %outPutPath% mkdir %outPutPath%

echo webdriver move

copy /Y %destFolderPath%\chromedriver-win64\chromedriver.exe %outPutPath%

rem zip�t�@�C����output�t�H���_���폜
del /q chromedriver-win64.zip
rmdir /q /s output


rem C:\ProgramData\scrapingFanbox��PATH��ʂ�
echo %PATH% | findstr "C:\ProgramData\scrapingFanbox" >NUL
if not ERRORLEVEL==1 (
    echo PATH has been set
) else (
    set PATH=%PATH%;C:\ProgramData\scrapingFanbox
)

rem scrapingFanbox.exe���N��
echo scrapingFanbox boot
call .\exe\scrapingFanbox.exe

echo �������I�����܂���
pause

