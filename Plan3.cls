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

Sub Init()

HostFolder = "C:\Users\Felipe\Desktop\ReplaceTest\"

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
            
                ws.Cells.SpecialCells(xlCellTypeConstants).Replace What:="hash", Replacement:="YAAA345454A", LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
                    
                
            Next ws
            
        End With
    Next
End Sub

