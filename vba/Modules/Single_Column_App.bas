Attribute VB_Name = "Single_Column_App"
Option Explicit

Public Sub Run_SingleColumn()
    Dim copyFromRange As CopyRange, _
        copyToRange As CopyRange, _
        i As Integer, _
        j As Integer, _
        currentItem As CopyItem, _
        CopyRanges As Collection, _
        CopyToRanges As Collection
    
    'turn off screen updating
    Application.ScreenUpdating = False
    On Error GoTo here
    
    'open copy ranges
    Set CopyRanges = Single_Column_CopyRanges
    Set CopyToRanges = Single_Column_CopyToRanges
    
    'copy items
    For j = 1 To CopyRanges.Count
        Set copyFromRange = CopyRanges(j)
        Set copyToRange = CopyToRanges(j)
        
        For i = 1 To copyFromRange.Count
            Set currentItem = copyFromRange.CopyItems(i)
            copyToRange.NextRowInRange.Value = currentItem.ItemsToCopy(1)
        Next i
        
    Next j
    
    'clean up
    copyFromRange.ItemWorkbook.Close
    Application.ScreenUpdating = False
    Exit Sub
here:
    
    'error handle
    copyFromRange.ItemWorkbook.Close
    Application.ScreenUpdating = False
    MsgBox Err.Description
End Sub
