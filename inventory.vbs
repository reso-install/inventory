Option Explicit

Dim ReportFile, strComputer, objWMIService, colItems, objItem, FS, File, Line
Dim ComputerName, Motherboard, Processor, Architecture, RAM, HDD, Display, OS

ReportFile  = "C:\InventoryReport.csv"
strComputer = "." 

' Win32_ComputerSystem
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystem",,48) 
For Each objItem in colItems 
    ComputerName = objItem.Name
    Motherboard = objItem.Manufacturer & " " & objItem.Model
    Architecture = objItem.SystemType
Next
ComputerName = """" & ComputerName & """"
Motherboard = """" & Motherboard & """"
Architecture = """" & Architecture & """"

' Win32_Processor
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_Processor",,48) 
For Each objItem in colItems 
    Processor = objItem.Name
Next
Processor = """" & Processor & """"

' Win32_PhysicalMemory
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_PhysicalMemory",,48) 
For Each objItem in colItems 
    RAM = RAM + Round(objItem.Capacity / (1024*1024*1024), 2)
Next
RAM = """" & RAM & "GB"""

' Win32_LogicalDisk
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_LogicalDisk where DriveType=3",,48) 
For Each objItem in colItems 
    HDD = HDD + Round(objItem.Size / (1024*1024*1024), 2)
Next
HDD = """" & HDD & "GB"""

' Win32_DesktopMonitor
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_DesktopMonitor",,48) 
For Each objItem in colItems
    If isNull(objItem.ScreenWidth) Or isNull(objItem.ScreenHeight) Then
        Display = objItem.Description
    Else
        Display = objItem.Description & " (" & objItem.ScreenWidth & "x" & objItem.ScreenHeight & ")"
    End If
Next
Display = """" & Display & """"

' Win32_OperatingSystem
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_OperatingSystem",,48) 
For Each objItem in colItems 
    OS = objItem.Caption
Next
OS = """" & OS & """"

Line = ComputerName & "," & Motherboard & "," & Processor & "," & RAM & "," & HDD & "," & Display & "," & OS & "," & Architecture
Line = CutSpaces(Line)
'Wscript.Echo Line

' Write to file
Set FS = CreateObject("Scripting.FileSystemObject")
If Not FS.FileExists(ReportFile) then
    Set File = FS.CreateTextFile(ReportFile, False)
    File.Write "Asset Tag,Motherboard,CPU,RAM,HDD,Display,OS,Architecture" & vbCrLf
Else
    ' Search for existing string
    Set File = FS.OpenTextFile(ReportFile)
    If InStr(File.ReadAll, Line) > 0 Then
        File.Close
        Wscript.Sleep 100
        WScript.Quit 1
    End If

    ' Reopen for appending
    File.Close
    Wscript.Sleep 100
    Set File = FS.OpenTextFile(ReportFile, 8, True) ' 8 - appending
End If

' Write to file
File.Write Line & vbCrLf
File.Close
WScript.Quit 0

Function CutSpaces (Input)
    Dim objRegEx
    Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Global = True
    objRegEx.Pattern = "\s+|\t+/ig"
    CutSpaces = objRegEx.Replace(Input, " ")
End Function