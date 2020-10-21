# CleanServer
Powershell script to clean Windows (Mail) Server

Basic and simple script to clean Windows (mail) Server. <br />
Do not run this when moving mailboxes or migrating! 

Remove files older than: 

    IIS (C:\Inetpub): 10 days
    Exchange (C:\ or E:\) : 30 days
    Windows - (Temp/Downloads): 10 days


Actions stored in: C:\log\*
