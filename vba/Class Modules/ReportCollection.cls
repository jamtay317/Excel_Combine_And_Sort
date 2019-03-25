VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public items As New Collection

Public Function AddOrUpdate(cItem As CopyItem) As ReportItem
    Dim item As ReportItem
    
    For Each item In items
        If item.RowDate = cItem.CopyItemDate Then
            UpdateItem item, cItem
            Set AddOrUpdate = item
            Exit Function
        End If
    Next item
    
    'we will never get here if already in Items so we need to create a new item
    Set item = New ReportItem
    item.RowDate = cItem.CopyItemDate
    UpdateItem item, cItem
    
    items.Add item
    Set AddOrUpdate = item
End Function

Public Sub AddCopyRange(cRange As CopyRange, myReportItem As ReportItem)
    
    If cRange.IsOffsetRange Then
        Set myReportItem.DownLinkCopyRange = cRange
    Else
        Set myReportItem.DCECopyRange = cRange
    End If
    
End Sub

Public Sub sort()
    Dim i As Integer, j As Integer, currentItem As ReportItem, compareItem As ReportItem, itemToMove As ReportItem, newItems As New Collection
    
    CollectionsSort.SortCollection items, "RowDate", True
    
End Sub

Private Sub UpdateItem(ByRef item As ReportItem, cItem As CopyItem)
    With item
        If cItem.IsOffset Then
            .downlinkItem.Add cItem
        Else
            .dceItem.Add cItem
        End If
    End With
End Sub
