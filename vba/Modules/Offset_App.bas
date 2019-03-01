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
            report.AddOrUpdate cItem
        Next cItem
    Next cRange
    
    'sort collection
    report.Sort
    
    'copy items in correct order and correct location
    i = copyToRange.StartRow + 1
    For Each rItem In report.items
        Set currentRange = copyToRange.NextRow
        
        CopyItemIntoWorkbook currentRange, rItem.DCEItem
        CopyItemIntoWorkbook currentRange.Offset(0, Offset_Width), rItem.DownLinkItem
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

Private Sub CopyItemIntoWorkbook(cRange As Range, cItem As CopyItem)
    Dim i As Integer
    
    If cItem Is Nothing Then Exit Sub
    
        cRange.Cells(1, 1).Value = cItem.CopyItemDate
    For i = 1 To cItem.ItemsToCopy.Count
            cRange.Cells(1, i + 1).Value = cItem.ItemsToCopy(i)
    Next i
End Sub

Private Sub CloseCopyFromWorkbooks(rangeCollection As Collection)
    Dim cRange As CopyRange
    
    For Each cRange In rangeCollection
        If cRange.IsOpen Then cRange.ItemWorkbook.Close
    Next cRange
End Sub
