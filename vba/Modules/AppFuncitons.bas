Attribute VB_Name = "AppFuncitons"
Option Explicit

Public Function GetCopyFromRanges() As Collection
    Dim myCollection As New Collection
        myCollection.Add Add_DCEs
        myCollection.Add Add_Downlinks
    Set GetCopyFromRanges = myCollection
End Function
Public Function GetCopyToRange() As CopyRange
    Dim copyToRange As New CopyRange
    
    Set GetCopyToRange = Copy_To_Range
End Function

Public Function OpenWorkbook(address As String) As Workbook
    Set OpenWorkbook = Application.Workbooks.Open(address, ReadOnly:=True)
End Function

Public Function IsWorkbookOpen(WorkbookName As String) As Boolean
    Dim wb As Workbook
    
    For Each wb In Application.Workbooks
        If wb.Name = WorkbookName Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
    
    IsWorkbookOpen = False
End Function

Public Function LastRow(startRange As Range) As Integer
    Dim startAddress As String: startAddress = startRange.address
    Dim columnNumber As Integer: columnNumber = startRange.Column
    LastRow = startRange.Worksheet.Cells(Rows.Count, columnNumber).End(xlUp).Row
End Function

Public Function LastUsedRow(ws As Worksheet) As Integer
   LastUsedRow = ws.UsedRange.Cells(ws.UsedRange.Rows.Count, 1).Row
End Function
