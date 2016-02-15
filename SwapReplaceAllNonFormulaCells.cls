'Replaces StringA with StringB and vice versa, Swapping the two.
'This will not be done with formulas
'This will be done in all excel files in HostFolder and all its sub folders

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim FileSystem As Object
Dim HostFolder As String
Dim StringA As String
Dim StringB As String

Sub Init()

HostFolder = "C:\Users\Felipe\Desktop\ReplaceTest\"
StringA = "one"
StringB = "two"

Set FileSystem = CreateObject("Scripting.FileSystemObject")
DoFolder FileSystem.GetFolder(HostFolder)

End Sub

Sub DoFolder(Folder)
    Dim SubFolder
    
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
    Next
    
    Dim File
    
    For Each File In Folder.Files
        With Workbooks.Open(File)
        
            Dim ws As Worksheet
            For Each ws In ActiveWorkbook.Worksheets
            
                On Error Resume Next
                ws.Cells.SpecialCells(xlCellTypeConstants).Replace What:=StringA, Replacement:="_AUXTEMPREPL_", LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
                On Error GoTo 0
                
                On Error Resume Next
                ws.Cells.SpecialCells(xlCellTypeConstants).Replace What:=StringB, Replacement:=StringA, LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
                On Error GoTo 0
                
                On Error Resume Next
                ws.Cells.SpecialCells(xlCellTypeConstants).Replace What:="_AUXTEMPREPL_", Replacement:=StringB, LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
                On Error GoTo 0
                
                
            Next ws
            
        End With
    Next
End Sub

