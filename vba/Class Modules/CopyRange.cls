VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public startAddress As String
Public WorkbookAddress As String
Public SheetName As String
Public IsOffsetRange As Boolean 'if isOffserRange is true this this is a downlink item
Public IsSingleColumn As Boolean

Public Property Get EndRow() As Integer
    EndRow = StartAddressRow + Count
End Property

Public Property Get StartAddressRow()
    StartAddressRow = Sheet.Range(startAddress).Row
End Property

Public Property Get StartRow() As Integer
    Dim startRange As Range
    Set startRange = Sheet.Range(startAddress)
    StartRow = startRange.Row
End Property

Private Property Get CopyWidth() As Integer
    If IsSingleColumn Then
        CopyWidth = 1
    Else
        CopyWidth = Copy_Width
    End If
End Property

Public Property Get CopyToColumns() As Collection
    Dim c As New Collection
    Dim splitItems As Variant, s As Variant
    
    If IsOffsetRange Then
        splitItems = Split(DownLink_CopyTo_Columns, ",")
    Else
        splitItems = Split(DCE_CopyTo_Columns, ",")
    End If
    
    For Each s In splitItems
        c.Add s
    Next s
    
    Set CopyToColumns = c
End Property

Public Property Get CopyToDateColumn() As String
    CopyToDateColumn = CopyToColumns(1)
End Property


Public Function ColumnsToCopy() As Collection
    Dim splitItems As Variant, s As Variant
    Set ColumnsToCopy = New Collection
    If IsOffsetRange Then
        splitItems = Split(DownLink_Columns, ",")
    Else
        splitItems = Split(DCE_Columns, ",")
    End If
    
    For Each s In splitItems
        ColumnsToCopy.Add s
    Next s
End Function

Public Property Get Range() As Range
   Dim startRange As Range
   Dim lr As Integer
   Dim startRowOffset As Integer
   
   
   Set startRange = Sheet.Range(startAddress)
   startRowOffset = startRange.Row
   lr = LastRow(startRange) - startRowOffset
   
   Set Range = Sheet.UsedRange
End Property

Public Property Get NextRow() As Range
   Dim startRange As Range, lr As Integer, startRowOffset As Integer, Cell1 As Range, Cell2 As Range

   
   Set startRange = Sheet.Range(startAddress)
   startRowOffset = startRange.Row
   lr = LastUsedRow(Sheet) - startRowOffset + 1
   
   Set Cell1 = startRange.Offset(lr, 0)
   Set Cell2 = startRange.Offset(lr, CopyWidth - 1)

   
   Set NextRow = Sheet.Range(Cell1, Cell2)
End Property

Public Property Get NextRowInRange() As Range
    Dim startRange As Range, _
        lRow As Integer, _
        nextCellAddress As String
    
    Set startRange = Sheet.Range(startAddress)
    lRow = LastRow(startRange)
    nextCellAddress = Sheet.Cells(lRow + 1, startRange.Column).address
    
    Set NextRowInRange = Sheet.Range(nextCellAddress)
    
End Property

Public Property Get Sheet() As Worksheet
    Set Sheet = ItemWorkbook.Sheets(SheetName)
End Property

Public Property Get Count() As Integer
    Dim lr As Integer, sr As Integer, startRange As Range

    Set startRange = Sheet.Range(startAddress)
    lr = LastRow(startRange)
    sr = startRange.Row
    Count = lr - sr + 1
End Property

Public Property Get CopyItems() As Collection
    Dim items As New Collection, cItem As CopyItem, i As Integer, _
           j As Integer, columnLetter As String, currentRow As Integer, _
           columnNumber As Integer
    
    For i = 1 To Count
        Set cItem = New CopyItem
        currentRow = StartRow + i - 1
        With cItem
            If IsSingleColumn Then
                columnNumber = Sheet.Range(startAddress).Column
                
                cItem.ItemsToCopy.Add Sheet.Cells(currentRow, columnNumber).Value
            Else
                
                'Sort column is always first item in list
                .CopyItemDate = Sheet.Range(SortColumn & currentRow).Value
                .IsOffset = IsOffsetRange
                For j = 1 To ColumnsToCopy.Count
                    columnLetter = ColumnsToCopy(j)
                    
                    .ItemsToCopy.Add Sheet.Range(columnLetter & currentRow).Value
                Next j
            End If
        End With
        items.Add cItem
        Debug.Print cItem.CopyItemDate, items.Count
    Next i
    
    Set CopyItems = items
End Property


Public Property Get SortColumn() As String
    SortColumn = ColumnsToCopy(1)
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = IsWorkbookOpen(WorkbookName)
End Property

Public Property Get WorkbookName() As String
    Dim paths() As String, length As Integer
    
    paths = Split(WorkbookAddress, "\")
    length = UBound(paths) - LBound(paths) + 1
    
    WorkbookName = paths(length - 1)
End Property

Public Property Get ItemWorkbook() As Workbook
    Dim wb As Workbook
    If Not IsOpen Then
        Set wb = OpenWorkbook(WorkbookAddress)
    Else
        Set wb = Application.Workbooks(WorkbookName)
    End If
    Set ItemWorkbook = wb
End Property
