Imports System.IO
Imports System.Web.Mail
Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public processpathinfo As String = ""
    Public interval As String = ""
    Public email As String = ""
    Public mailserver As String = ""

    Public result As String = ""

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal filesparsed As Long, ByVal entriesfound As Long, ByVal currentfile As String, ByVal totalfiles As Long)
    Public Event WorkerFileUpdate(ByVal entriesfound As Long, ByVal currentfile As String)


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try
            If message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Worker Thread encountered the following problem: " & vbCrLf & message)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & message)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch ex As Exception
            MsgBox("An error occurred in IIS W3SVC1 Log Monitor's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try
            Dim path, path2 As String
            path = (Application.StartupPath & "\").Replace("\\", "\") & "IIS W3SVC1 Log Monitor.mdb"
            path2 = (Application.StartupPath & "\").Replace("\\", "\") & "IIS W3SVC1 Log Template.mdb"
            Dim databasebackup As FileInfo = New FileInfo(path2)
            If databasebackup.Exists = True Then
                Dim database As FileInfo = New FileInfo(path)
                If database.Exists = False Then
                    If databasebackup.Exists = True Then
                        File.Copy(path2, path)
                    End If
                End If
                Dim dinfo As DirectoryInfo = New DirectoryInfo(processpathinfo)

                database = (New FileInfo(path))
                While database.Exists = False

                End While

                If database.Exists = True Then
                    If dinfo.Exists = True Then
                        If dinfo.GetFiles.Length > 0 Then

                            Dim Conn As Data.OleDb.OleDbConnection
                            Conn = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & database.FullName & ";")
                            Conn.Open()


                            dinfo = New DirectoryInfo(processpathinfo)
                            Dim finfo As FileInfo
                            Dim count As Integer
                            count = dinfo.GetFiles.Length
                            Dim runner As Long = 0
                            Dim linecount As Long
                            linecount = 0
                            Dim filelinecount As Long
                            filelinecount = 0
                            Dim totalfilecount As Long = dinfo.GetFiles.Length
                            Dim run As Integer
                            Dim table(22) As String

                            table(0) = "log_date"
                            table(1) = "log_time"
                            table(2) = "log_s_sitename"
                            table(3) = "log_s_computername"
                            table(4) = "log_s_ip"
                            table(5) = "log_cs_method"
                            table(6) = "log_cs_uri_stem"
                            table(7) = "log_cs_uri_query"
                            table(8) = "log_s_port"
                            table(9) = "log_cs_username"
                            table(10) = "log_c_ip"
                            table(11) = "log_cs_version"
                            table(12) = "log_cs_User_Agent"
                            table(13) = "log_cs_Cookie"
                            table(14) = "log_cs_Referer"
                            table(15) = "log_cs_host"
                            table(16) = "log_sc_status"
                            table(17) = "log_sc_substatus"
                            table(18) = "log_sc_win32_status"
                            table(19) = "log_sc_bytes"
                            table(20) = "log_cs_bytes"
                            table(21) = "log_time_taken"


                            For Each finfo In dinfo.GetFiles
                                Try
                                    RaiseEvent WorkerProgress(runner, linecount, finfo.Name, dinfo.GetFiles.Length)
                                    runner = runner + 1
                                    If runner < count Then
          


                                        Dim filereader As StreamReader
                                        filereader = New StreamReader(finfo.FullName, True)




                                        Dim sql As String
                                        Dim lineread As String
                                        filelinecount = 0
                                        lineread = filereader.ReadLine()
                                        While Not lineread Is Nothing


                                            If Not lineread.StartsWith("#") = True And lineread.Length > 0 Then
                                                linecount = linecount + 1
                                                filelinecount = filelinecount + 1
                                                RaiseEvent WorkerProgress(runner, linecount, finfo.Name, dinfo.GetFiles.Length)
                                                RaiseEvent WorkerFileUpdate(filelinecount, finfo.Name)
                                                Dim results As String()
                                                lineread = lineread.Replace("""", "``")
                                                results = lineread.Split(" ")
                                                Dim small As String

                                                run = 0
                                                If results.Length = 22 Then


                                                    For Each small In results
                                                        Dim count2 As Long

                                                        sql = "select * from " & table(run) & " where log_value = """ & small & """"

                                                        Dim recset As Data.OleDb.OleDbCommand = New Data.OleDb.OleDbCommand(sql, Conn)
                                                        Dim recread As Data.OleDb.OleDbDataReader
                                                        recread = recset.ExecuteReader()
                                                        Dim action As Boolean
                                                        If recread.HasRows Then
                                                            recread.Read()
                                                            count2 = CLng(recread.Item(recread.GetOrdinal("log_count"))) + 1
                                                            action = True
                                                        Else
                                                            count2 = 1
                                                            action = False
                                                        End If

                                                        recread.Close()
                                                        recset.Dispose()

                                                        If small.Length > 255 Then
                                                            small = small.Substring(0, 255)
                                                        End If
                                                        If action = True Then
                                                            sql = "update " & table(run) & " set log_count = " & count2 & " where log_value = """ & small & """"
                                                        Else
                                                            sql = "insert into " & table(run) & " values (""" & small & """," & count2 & ")"
                                                        End If

                                                        recset = New Data.OleDb.OleDbCommand(sql, Conn)
                                                        Dim result As Integer
                                                        result = (recset.ExecuteNonQuery())
                                                        recset.Dispose()
                                                        run = run + 1

                                                    Next
                                                Else

                                                    sql = "insert into log_errors values (""" & lineread & """)"
                                                    Dim recset As Data.OleDb.OleDbCommand
                                                    recset = New Data.OleDb.OleDbCommand(sql, Conn)
                                                    Dim result As Integer
                                                    result = (recset.ExecuteNonQuery())
                                                    recset.Dispose()

                                                End If
                                            End If


                                            lineread = filereader.ReadLine()
                                        End While

                                        filereader.Close()
                                        File.Delete(finfo.FullName)
                                    End If
                                Catch ex As Exception
                                    Error_Handler("An """ & ex.Message & """ error occurred while parsing the log file.")
                                End Try

                            Next






                            Conn.Close()
                            Conn.Dispose()
                            database = Nothing


                            '        Dim body As String

                            '        body = "The following applications have been detected as not currently running. Application restarts have been attempted, but please dispatch or inform the necessary maintenance staff to ensure system availability."
                            '        body = body & vbCrLf & vbCrLf & "******************************" & vbCrLf & vbCrLf
                            '        error_message = error_message.Replace(Chr(13), " ")
                            '        body = body & error_message.Trim
                            '        body = body & vbCrLf & vbCrLf & "******************************" & vbCrLf & vbCrLf & "This is an auto-generated email submitted from IIS W3SVC1 Log Monitor 1.0 at " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & ", running on:"
                            '        body = body & vbCrLf & vbCrLf & "Machine Name: " + Environment.MachineName
                            '        body = body & vbCrLf & "OS Version: " & Environment.OSVersion.ToString()
                            '        body = body & vbCrLf & "User Name: " + Environment.UserName

                            '        TextMail(mailserver, "applications_monitor@commerce.uct.ac.za", email, "Critical Applications Restart Required", body)
                            '    Catch ex As Exception
                            '        Error_Handler("An """ & ex.Message & """ error occurred while sending a notification email. The program will attempt to recover shortly.")
                            '    End Try
                        Else
                            RaiseEvent WorkerProgress(0, 0, "", 0)
                            Error_Handler("The specified log folder is empty.")
                        End If
                    Else
                        RaiseEvent WorkerProgress(0, 0, "", 0)
                        Error_Handler("The specified log folder doesn't exist.")
                    End If
                Else
                    RaiseEvent WorkerProgress(0, 0, "", 0)
                    Error_Handler("No valid Database could be located to store results in.")
                End If
            Else
                RaiseEvent WorkerProgress(0, 0, "", 0)
                Error_Handler("No valid Database could be located to store results in.")
            End If
            result = "Success"
            RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            RaiseEvent WorkerComplete(result)
        End Try

        WorkerThread.Abort()
    End Sub

    Private Sub ApplicationLauncher(ByVal apptoRun As String)
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = apptoRun
            myProcess.StartInfo.UseShellExecute = True

            myProcess.StartInfo.CreateNoWindow = False


            myProcess.StartInfo.RedirectStandardInput = False
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.StartInfo.RedirectStandardError = False
            myProcess.Start()


        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching External Application. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Function DosShellCommand(ByVal AppToRun As String, Optional ByVal silent As Boolean = True) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            If silent = True Then
                myProcess.StartInfo.CreateNoWindow = True
            Else
                myProcess.StartInfo.CreateNoWindow = False
            End If

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching DOS shell. The program will attempt to recover shortly.")
        End Try
        Return s
    End Function

    Public Function TextMail(ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, Optional ByRef strErrMsg As String = "") As Boolean
        Dim objMail As MailMessage

        Try
            Dim emailaddys As String()
            emailaddys = strTo.Split(";")

            Dim counter As Integer = 0
            For counter = 0 To emailaddys.Length - 1
                objMail = New MailMessage
                objMail.BodyFormat = MailFormat.Text
                objMail.From = "application_monitor@commerce.uct.ac.za"
                objMail.To = emailaddys(counter).Trim
                objMail.Subject = strSubj
                objMail.Body = strBody

                SmtpMail.SmtpServer = mailserver
                SmtpMail.Send(objMail)
            Next
            TextMail = True

        Catch ex As Exception
            TextMail = False
            Error_Handler("An """ & ex.Message & """ error occurred while sending the Error Alert email. The program will attempt to recover shortly.")
        End Try
    End Function

    Public Function TextMail(ByVal SmtpServer As String, ByVal strFrom As String, ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, Optional ByRef strErrMsg As String = "") As Boolean
        Dim objMail As MailMessage

        Try
            Dim emailaddys As String()
            emailaddys = strTo.Split(";")

            Dim counter As Integer = 0
            For counter = 0 To emailaddys.Length - 1


                objMail = New MailMessage
                objMail.BodyFormat = MailFormat.Text
                objMail.From = strFrom
                objMail.To = emailaddys(counter).Trim
                objMail.Subject = strSubj
                objMail.Body = strBody

                SmtpMail.SmtpServer = SmtpServer
                SmtpMail.Send(objMail)
            Next
            TextMail = True

        Catch ex As Exception
            TextMail = False
            Error_Handler("An """ & ex.Message & """ error occurred while sending the Error Alert email. The program will attempt to recover shortly.")
        End Try
    End Function

End Class
