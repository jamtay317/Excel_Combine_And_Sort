Attribute VB_Name = "Setup_Singe_Column_App"
Option Explicit

Public Function Single_Column_CopyRanges() As Collection
    Dim Range1 As New CopyRange, Range2 As New CopyRange
    
    Set Single_Column_CopyRanges = New Collection
    
    With Range1
        .startAddress = "A2"
        .SheetName = "Sheet1"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\Data.xlsx"
        .IsSingleColumn = True
    End With
    Single_Column_CopyRanges.Add Range1
    
    With Range2
        .startAddress = "A2"
        .SheetName = "Sheet2"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\Data.xlsx"
        .IsSingleColumn = True
    End With
    Single_Column_CopyRanges.Add Range2
    
End Function

Public Function Single_Column_CopyToRanges() As Collection
    Set Single_Column_CopyToRanges = New Collection
    Dim Range1 As New CopyRange, Range2 As New CopyRange
    
    With Range1
        .startAddress = "A2"
        .SheetName = "Sheet1"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\CopyAndSort.xlsm"
        .IsSingleColumn = True
    End With
    Single_Column_CopyToRanges.Add Range1
    
    With Range2
        .startAddress = "C2"
        .SheetName = "Sheet1"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\CopyAndSort.xlsm"
        .IsSingleColumn = True
    End With
    Single_Column_CopyToRanges.Add Range2
End Function
