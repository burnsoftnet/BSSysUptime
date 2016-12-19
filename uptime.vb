Imports System
Imports Microsoft.Win32
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Windows.Forms
Public Module Uptime
    Dim bOnly As Boolean 'Display only uptime flag
    Dim bCopy As Boolean 'Copy uptime to clipboard flag
    Dim bNoDisp As Boolean 'No display flag
    Dim bNoEvent As Boolean 'Don't write to event log flag
    Dim sHost As String 'Remote machine to retrieve uptime
    Sub Init()
        Dim DidExist As Boolean = False

        sHost = GetCommand("host", System.Environment.MachineName,,)
        bOnly = GetCommand("only", True,,)      'Text Only
        bCopy = GetCommand("copy", False,,)     'Copy to Clipboard
        bNoDisp = GetCommand("nodisplay", False,,)  'Suppress Display
        bNoEvent = GetCommand("noevent", False,,)   'Suppress Write to Event Log
        If (GetCommand("help", False,,) Or GetCommand("?", False,,)) Then
            Call ShowHelp() 'Show Help File and Exit
        End If
    End Sub
    Sub ShowHelp()
        Console.WriteLine(vbNewLine & "BurnSoft System Uptime Tool")
        Console.WriteLine(String.Format("Version {0}", Application.ProductVersion.ToString))
        Console.WriteLine(My.Application.Info.ProductName & "  " & My.Application.Info.Copyright)
        Console.WriteLine()
        Console.WriteLine("uptime /host=[machine] /copy /only /nodisplay /noevent /help /?" & vbNewLine)
        Console.WriteLine(vbTab & "/host=[machine] - Optional is you want to check a remote machine, otherwise it will look at local.")
        Console.WriteLine(vbTab & "/copy - Copy uptime to clipboard")
        Console.WriteLine(vbTab & "/only - Show uptime only")
        Console.WriteLine(vbTab & "/nodisplay - Write nothing to console")
        Console.WriteLine(vbTab & "/noevent - Don't write event to NT Event Log")
        Console.WriteLine(vbTab & "/help, /? - Show help" & vbNewLine)
        Console.WriteLine()
        Call PressToExit()
    End Sub
    Sub PressToExit()
        Console.WriteLine()
        Console.WriteLine("Press enter/return key to continue.")
        Console.Read()
        Application.Exit()
    End Sub
    Public Sub Main(ByVal ParamArray args() As String)
        Try
            Dim oPerf As System.Diagnostics.PerformanceCounter
            Dim oUptime As New System.Diagnostics.CounterSample
            Dim dUptime As Double
            Dim dBuffer As Double
            Dim lUnitBuff As Long
            Dim sOutput As String
            Dim a As Long
            Dim MyAppPath As String = System.Environment.CurrentDirectory
            Dim MyAppName As String = System.Environment.CommandLine

            Call Init()
            Console.WriteLine("Host name is " & sHost)

            oPerf = New System.Diagnostics.PerformanceCounter
            oPerf.CategoryName = "System"
            oPerf.CounterName = "System Up Time"

            If Len(sHost) > 0 Then   'If remote machine was specified
                oPerf.MachineName = sHost
            End If
            oUptime = oPerf.NextSample
            dUptime = oUptime.Calculate(oUptime, oPerf.NextSample)
            '--Weeks---
            If dUptime >= 604800 Then
                dBuffer = dUptime / 604800
                lUnitBuff = CLng(Mid(CStr(dBuffer), 1, InStr(1, CStr(dBuffer), ".") - 1))

                If lUnitBuff > 1 Then
                    sOutput = sOutput & CStr(lUnitBuff) & " weeks, "
                Else
                    sOutput = sOutput & "1 week, "
                End If
                dUptime = dUptime - CDbl(lUnitBuff * 604800)
            End If
            '--Days----
            If dUptime >= 86400 Then
                dBuffer = dUptime / 86400
                lUnitBuff = CLng(Mid(CStr(dBuffer), 1, InStr(1, CStr(dBuffer), ".") - 1))
                If lUnitBuff > 1 Then
                    sOutput = sOutput & CStr(lUnitBuff) & " days, "
                Else
                    sOutput = sOutput & "1 day, "
                End If
                dUptime = dUptime - CDbl(lUnitBuff * 86400)
            End If

            '--Hours----
            If dUptime >= 3600 Then
                dBuffer = dUptime / 3600
                lUnitBuff = CLng(Mid(CStr(dBuffer), 1, InStr(1, CStr(dBuffer), ".") - 1))
                If lUnitBuff > 1 Then
                    sOutput = sOutput & CStr(lUnitBuff) & " hours, "
                Else
                    sOutput = sOutput & "1 hour, "
                End If
                dUptime = dUptime - CDbl(lUnitBuff * 3600)
            End If
            '--Minutes----
            If dUptime >= 60 Then
                dBuffer = dUptime / 60
                lUnitBuff = CLng(Mid(CStr(dBuffer), 1, InStr(1, CStr(dBuffer), ".") - 1))
                If lUnitBuff > 1 Then
                    sOutput = sOutput & CStr(lUnitBuff) & " minutes, "
                Else
                    sOutput = sOutput & "1 minute, "
                End If
                dUptime = dUptime - CDbl(lUnitBuff * 60)
            End If
            '--Seconds----
            If dUptime >= 1 Then
                dBuffer = dUptime
                lUnitBuff = CLng(Mid(CStr(dBuffer), 1, InStr(1, CStr(dBuffer), ".") - 1))
                If lUnitBuff > 1 Then
                    sOutput = sOutput & CStr(lUnitBuff) & " seconds"
                Else
                    sOutput = sOutput & "1 second"
                End If
            End If

            'Write it!
            If Not bNoDisp Then Call WriteToConsole(sOutput)
            'Record to NT event log
            If Not bNoEvent Then Call WriteEventLog(sOutput)
            'Copy to clipboard
            If bCopy Then Call CopyToClipBoard(sOutput)
        Catch ex As Exception
            'Write Error to screen
            Console.WriteLine(vbNewLine & " An Error Has Occured:" & vbNewLine)
            Console.WriteLine(" " & Err.Description & vbNewLine)
            If Not bNoEvent Then
                Dim logUptime As New EventLog
                logUptime.Log = "Application"
                logUptime.Source = "uptime"
                logUptime.WriteEntry("An Error Has Occured:" & vbNewLine & vbNewLine & Err.Description)
            End If
        End Try
    End Sub
    Sub CopyToClipBoard(sOutPut As String)
        If bOnly Then 'only uptime
            System.Windows.Forms.Clipboard.SetDataObject(sOutPut, True)
        Else
            System.Windows.Forms.Clipboard.SetDataObject(FormatOutput(sOutPut, sHost), True)
        End If
    End Sub
    Sub WriteEventLog(sOutPut As String)
        Dim logUptime As New EventLog
        logUptime.Log = "Application"
        logUptime.Source = "uptime"
        logUptime.WriteEntry("System uptime retrieved on " & Now & FormatOutput(sOutPut, sHost))
    End Sub
    Sub WriteToConsole(sOutput As String)
        If bOnly Then 'Only uptime
            Console.Write(sOutput)
            Call PressToExit()
        Else
            Console.WriteLine(FormatOutput(sOutput, sHost))
            Call PressToExit()
        End If
    End Sub
    Private Function FormatOutput(ByVal sUptime As String, Optional ByVal sHost As String = "") As String
        If Len(sHost) <= 0 Then sHost = Environment.MachineName
        Return vbNewLine & sHost & " Uptime: " & vbTab & sUptime & vbNewLine
    End Function


End Module
