IIS W3SVC1 Log Monitor
======================

IIS W3SVC1 Log Monitor is used to parse IIS W3SVC1 format log files and place the results into a MsAccess database. After it parses a file it removes it from the system, though it will not parse the last file present in the log folder as this would be the file still in use by IIS. By setting the monitoring interval, IIS W3SVC1 Log Monitor will work in tandem with IIS, consuming the log files as they are produced by IIS.

The current format of valid log files is as follows: (single line, space delimited)
date time s-sitename s-computername s-ip cs-method cs-uri-stem cs-uri-query 
s-port cs-username c-ip cs-version cs(User-Agent) cs(Cookie) cs(Referer)
cs-host sc-status sc-substatus sc-win32-status sc-bytes cs-bytes 
time-taken

Created by Craig Lotter, September 2005

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2003
Implements concepts such as threading, file manipulation, sql and timer programming.
Level of Complexity: simple
