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

Public Items As New Collection

Public Sub AddOrUpdate(cItem As CopyItem)
    Dim item As ReportItem
    
    For Each item In Items
        If item.RowDate = cItem.CopyItemDate Then
            UpdateItem item, cItem
            Exit Sub
        End If
    Next item
    
    'we will never get here if already in Items so we need to create a new item
    Set item = New ReportItem
    item.RowDate = cItem.CopyItemDate
    UpdateItem item, cItem
    
    Items.Add item
End Sub

Public Sub Sort()
    Dim i As Integer, j As Integer, currentItem As ReportItem, compareItem As ReportItem, itemToMove As ReportItem
    
    For i = 1 To Items.Count
        Set currentItem = Items(i)
        
        For j = i + 1 To Items.Count
            Set compareItem = Items(j)
            
            If currentItem.RowDate > compareItem.RowDate Then
                Set itemToMove = compareItem
                
                Items.Remove j
                
                Items.Add itemToMove, itemToMove.RowDate & "", i
            End If
        Next j
    Next i
End Sub

Private Sub UpdateItem(ByRef item As ReportItem, cItem As CopyItem)
    With item
        If cItem.IsOffset Then
            Set .DownLinkItem = cItem
        Else
            Set .DCEItem = cItem
        End If
    End With
End Sub
