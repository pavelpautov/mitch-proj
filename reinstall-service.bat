@echo OFF
echo Stopping old service version...
net stop "DocxToPdf"
echo Uninstalling old service version...
sc delete "DocxToPdf"

echo Installing service...
rem DO NOT remove the space after "binpath="!
sc create "DocxToPdf" binpath= "C:\Pavel\UpWork\Util-for-convert-docx-to-pdf\github-proj\DocxToPdfService\DocxToPdfService\bin\Debug\DocxToPdfService.exe" start= demand
echo Starting server complete
pause