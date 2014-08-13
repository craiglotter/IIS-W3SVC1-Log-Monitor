IIS-W3SVC1-Log-Monitor
======================

IIS W3SVC1 Log Monitor is used to parse IIS W3SVC1 format log files and place the results into a MsAccess database. After it parses a file it removes it from the system, though it will not parse the last file present in the log folder as this would be the file still in use by IIS. By setting the monitoring interval, IIS W3SVC1 Log Monitor will work in tandem with IIS, consuming the log files as they are produced by IIS. Created by Craig Lotter, September 2005
