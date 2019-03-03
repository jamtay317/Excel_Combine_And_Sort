Attribute VB_Name = "Offset_App"
Option Explicit

Public Sub Run_Offset()
    Dim rangeCollection As Collection, _
           ItemsToCopy As New Collection, _
           cRange As CopyRange, _
           cItem As CopyItem, _
           copyToRange As CopyRange, _
           report As New ReportCollection, _
           i As Integer, _
           rItem As ReportItem, _
           currentRange As Range
    
    Application.ScreenUpdating = False
    
    'Error Handle
    On Error GoTo here
    
    'setup
    Set rangeCollection = GetCopyFromRanges()
    Set copyToRange = GetCopyToRange
    
    'get items to copy
    For Each cRange In rangeCollection
        For Each cItem In cRange.CopyItems
            Set rItem = report.AddOrUpdate(cItem)
            report.AddCopyRange cRange, rItem
        Next cItem
    Next cRange
    
    'sort collection
    report.Sort
    
    'copy items in correct order and correct location
    i = copyToRange.StartRow + 1
    For Each rItem In report.items
        Set currentRange = copyToRange.NextRow
        
        CopyItemIntoWorkbook rItem, i, copyToRange.Sheet
        i = i + 1
    Next rItem
    
    
    'clean up
    CloseCopyFromWorkbooks rangeCollection
    Application.ScreenUpdating = True
    
    Exit Sub
here:
    CloseCopyFromWorkbooks rangeCollection
    Application.ScreenUpdating = True
    MsgBox Err.Description
End Sub

Private Sub CopyItemIntoWorkbook(rItem As ReportItem, rowNumber As Integer, ByRef copyToSheet As Worksheet)
    Dim letter As String, i As Integer
    
    'this should only try to copy the value to the new sheet if there is a copy range
    If Not rItem.DCECopyRange Is Nothing Then
        For i = 1 To rItem.DCECopyRange.CopyToColumns.Count
            letter = rItem.DCECopyRange.CopyToColumns(i)
            copyToSheet.Range(letter & rowNumber).Value = rItem.DCEItem.ItemsToCopy(i)
        Next i
    End If
    
    If Not rItem.DownLinkCopyRange Is Nothing Then
        For i = 1 To rItem.DownLinkCopyRange.CopyToColumns.Count
            letter = rItem.DownLinkCopyRange.CopyToColumns(i)
            copyToSheet.Range(letter & rowNumber).Value = rItem.DownLinkItem.ItemsToCopy(i)
        Next i
    End If
    
End Sub

Private Sub CloseCopyFromWorkbooks(rangeCollection As Collection)
    Dim cRange As CopyRange
    
    For Each cRange In rangeCollection
        If cRange.IsOpen Then cRange.ItemWorkbook.Close
    Next cRange
End Sub
