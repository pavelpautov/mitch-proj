# Docx to Pdf windows service

Windows service that automatically transforms *.docx files to PDF. Written on .Net Framework. 
Word .docx files operated using Microsoft Office primary interop assemblies (PIA) - so need to have Office/Word installed. 

Source directory - for .docx files. Hardcoded in `Service1.DocxFolderPath`

Destination directory - for generated PDFs. Hardcoded in `Service1.PdfForderPath`


# How this works:

Service monitor 'Source' directory (see Serivce1.cs - used FileSystemWatcher class).
If any *.docx files will be copied/created in 'source' directory - the event is fired (see `Service1.OnStart`)

In `OnChanged` we check if we already have such PDF - we do generation only if there is no PDF already created.

Detection of files and any exceptions are logged in txt file - you could find it in subfolder `Logs` in the directory where .exe file of service situated. The log file name is given from the current date - for example, `ServiceLog_21.03.2020.txt`


# How to install on PC

First, you need to open project in VS and rebuild it. Say you have .exe file in place (my example)
`C:\Pavel\UpWork\Util-for-convert-docx-to-pdf\github-proj\DocxToPdfService\DocxToPdfService\bin\Debug\DocxToPdfService.exe`

Exe file could not be started in Windows because this is a service.

The simplest way to install - use reinstall-service.bat file in this GitHub repo. - see inside: there are 3 commands. First, stop windows service (if exist), then delete service, and then install service.

You need to run batch file 'As Admin'. After the batch is successfully executed, you could find DocxToPdf service in Windows Service Manager (services.msc) - need to start the service. - that's all. Now it works.

You could copy docx files in 'source' directory and check what happens.

See demo video ...(don't forget to add a link there)