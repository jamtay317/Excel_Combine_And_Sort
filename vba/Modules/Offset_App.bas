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
    
    'Application.ScreenUpdating = False
    
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
    Call report.Sort
    
    'copy items in correct order and correct location
    i = copyToRange.StartRow
    For Each rItem In report.items
        If rItem.RowDate > 0 Then
            Set currentRange = copyToRange.nextRow
            i = GetNextCopyRow(copyToRange.Sheet, rItem)
        
            CopyItemIntoWorkbook rItem, i, copyToRange.Sheet
            
        End If
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

Private Function CopyItemIntoWorkbook(rItem As ReportItem, rowNumber As Integer, ByRef copyToSheet As Worksheet) As Boolean
    Dim letter As String, i As Integer, dItem As CopyItem, nextRow As Integer
    
    'this should only try to copy the value to the new sheet if there is a copy range
    If Not rItem.DCECopyRange Is Nothing Then
        nextRow = rowNumber
        For Each dItem In rItem.dceItem
            For i = 1 To rItem.DCECopyRange.CopyToColumns.Count
                letter = rItem.DCECopyRange.CopyToColumns(i)
           
                'rItem.DCEItem should be a collection of rows that need coppied
                
                copyToSheet.Range(letter & nextRow).Value = dItem.ItemsToCopy(i)
           
            Next i
            
            nextRow = nextRow + 1

        Next dItem
    End If
    
    If Not rItem.DownLinkCopyRange Is Nothing Then
        nextRow = rowNumber
        For Each dItem In rItem.downlinkItem
            
            For i = 1 To rItem.DownLinkCopyRange.CopyToColumns.Count
                letter = rItem.DownLinkCopyRange.CopyToColumns(i)
    
                'rItem.DownLinkItem should be a collection of rows that need coppied
                copyToSheet.Range(letter & rowNumber).Value = dItem.ItemsToCopy(i)
            Next i
            nextRow = nextRow + 1
        Next dItem
    End If
    
End Function

Private Sub CloseCopyFromWorkbooks(rangeCollection As Collection)
    Dim cRange As CopyRange
    
    For Each cRange In rangeCollection
        If cRange.IsOpen Then cRange.ItemWorkbook.Close
    Next cRange
End Sub


Private Function GetNextCopyRow(ws As Worksheet, rItem As ReportItem) As Integer
    Dim columnNumber As Integer, dceNextRow As Integer, downLinkRow As Integer
    
    If Not IsNull(DCEDateColumn) And Not DCEDateColumn = "" Then
        columnNumber = ws.Range(DCEDateColumn & 1).Column
        dceNextRow = ws.Cells(Rows.Count, columnNumber).End(xlUp).Row + 1
    End If
    
    If Not IsNull(DownLinkDateColumn) And Not DownLinkDateColumn = "" Then
        columnNumber = ws.Range(DownLinkDateColumn & 1).Column
        downLinkRow = ws.Cells(Rows.Count, columnNumber).End(xlUp).Row + 1
    End If
    
    If downLinkRow > dceNextRow Then
        GetNextCopyRow = downLinkRow
    Else
        GetNextCopyRow = dceNextRow
    End If
    
    
    
End Function
