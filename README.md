MailSorter
==========
For "Inbox Zero" enthusiasts: A simple tool for Outlook, which allows you to click one button to mark a message as read and sort it into a folder that matches the sender's name.

To install, build the project and double-click the .vsto file in the bin\debug folder (ensure Outlook is closed).  When you open Outlook, you should see a "Mail Sorter" section in your ribbon.

Known Issues
======
If you get an error about VSTOInstaller.exe.config not being found, rename C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe.config to VSTOInstaller.exe.config.bak and run the bin\debug\*.vsto install again.  You can undo the rename once you're done.
