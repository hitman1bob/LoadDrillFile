Imports System.IO
Imports System.Diagnostics
Imports System.Configuration
Imports CAMMaster
Module Module1

    Sub Main()
        '--------------------------------------------------
        'Determine the number of CAMMaster processes
        '--------------------------------------------------
        Dim strWindowTitle As String
        Dim plist() As Process = Process.GetProcessesByName("CAMMAS~1")

        If plist.Length > 2 Then
            'There is more than one procewss with this name.
            'Therefore, this program should end.
            Call printToConsole("More than one CAMMaster is running!  I will exit now.")
            End
        End If

        '--------------------------------------------------
        'Determine the tool number and get the files
        '--------------------------------------------------
        Dim DrillFile, ImportDrill As String
        Dim btimportdrill As New CAMMaster.Tool
        Dim ToolNo As String = ""

        For Each p As Process In Process.GetProcesses()
            If Not (p.MainWindowTitle Is Nothing) Then
                If p.ProcessName = "CAMMAS~1" Then
                    strWindowTitle = p.MainWindowTitle
                    ToolNo = Left(strWindowTitle, 5)
                    'printToConsole("CAMMaster was found and is titled " & strWindowTitle)
                    'printToConsole("The tool number is " & Left(strWindowTitle, 5))
                    'Console.Write("Press any key to continue...")
                    'Console.ReadKey()
                End If
            End If
        Next

        Dim UFPath As String = My.Settings("UneditedFilesPath")
        Dim EFPath As String = My.Settings("EditedFilesPath")
        Dim DrillExtension As String = My.Settings("DrillExtension")
        Dim DrillReportExtension As String = My.Settings("DrillReportExtension")
        Dim RoutReportExtension As String = My.Settings("RoutReportExtension")

        DrillFile = (ToolNo & DrillExtension)
        Dim hll As Int16 = btimportdrill.GetHighestLoadedLayer + 1

        btimportdrill.CurrentLayer = hll

        If System.IO.File.Exists(UFPath + DrillFile) = True Then
            ' Import it into CAMMaster
            ImportDrill = (UFPath + DrillFile)
            btimportdrill.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
            btimportdrill.AddToFileList(ImportDrill)
            btimportdrill.Import("Drill")
        ElseIf System.IO.File.Exists(EFPath & DrillFile) = True Then    ' Check the second path
            'Import it into CAMMaster
            ImportDrill = (EFPath & DrillFile)
            btimportdrill.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
            btimportdrill.AddToFileList(ImportDrill)
            btimportdrill.Import("Drill")
        Else
            'Give error message
            printToConsole("The drill file was not found.")
        End If

    End Sub
    Public Function FileExists(ByVal FileFullPath As String) As Boolean

        Dim f As New IO.FileInfo(FileFullPath)
        Return f.Exists

    End Function

    Public Function FolderExists(ByVal FolderPath As String) As Boolean

        Dim f As New IO.DirectoryInfo(FolderPath)
        Return f.Exists

    End Function

    Public Sub KillWindow(ByVal strWindowTitle As String)

        Dim plist As Process() = Process.GetProcesses
        For Each p As Process In plist
            If Not (p.MainWindowTitle Is Nothing) Then
                If p.MainWindowTitle = strWindowTitle Then
                    p.Kill()
                End If
            End If
        Next

    End Sub

    Sub printToConsole(ByVal anyString As String)
        Console.WriteLine(anyString)
    End Sub
End Module
