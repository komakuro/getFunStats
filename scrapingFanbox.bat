@echo off
cd /d %~dp0

rem chocolateyがなければWgetとあわせてインストールする
where /q choco
if %errorlevel% == 0 ( :: コマンドが存在すれば
    echo choco exists
) else (               :: コマンドが存在しなければ
	echo Y|"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
    echo Y|choco install Wget
)

where /q wget
if %errorlevel% == 0 ( :: コマンドが存在すれば
    echo wget exists
) else (               :: コマンドが存在しなければ
	echo Y|choco install Wget
)

rem Chromeのバージョンを確認して、同じバージョンのChromeDriverをダウンロード
echo webdriver download

rem set version=dir /B /AD "C:\Program Files\Google\Chrome\Application"
set versionPath="C:\Program Files\Google\Chrome\Application"
for /f "usebackq" %%A in (`dir /B /AD %versionPath%`) do set version=%%A&goto :exit_for
:exit_for
echo version is %version%
set url=https://storage.googleapis.com/chrome-for-testing-public/%version%/win64/chromedriver-win64.zip
wget %url%


rem ダウンロードしたzipファイルを解凍
echo unzip download file
set zipFilePath=.\chromedriver-win64.zip
set destFolderPath=.\output

set psCommand=powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Path %zipFilePath% -DestinationPath %destFolderPath% -Force
%psCommand%


rem C:\ProgramData以下にscrapingFanboxフォルダを作り、ChromeDriverをコピー
echo webdriver move stanby

set outPutPath=C:\ProgramData\scrapingFanbox
if not exist %outPutPath% mkdir %outPutPath%

echo webdriver move

copy /Y %destFolderPath%\chromedriver-win64\chromedriver.exe %outPutPath%


rem zipファイルとoutputフォルダを削除
del /q chromedriver-win64.zip
rmdir /q /s output


rem C:\ProgramData\scrapingFanboxにPATHを通す
echo %PATH% | findstr "C:\ProgramData\scrapingFanbox" >NUL
if not ERRORLEVEL==1 (
    echo PATH has been set
) else (
    set PATH=%PATH%;C:\ProgramData\scrapingFanbox
)
pause
