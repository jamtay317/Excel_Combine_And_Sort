Attribute VB_Name = "Setup_Offset_App"
Option Explicit

Public Const Copy_Width As Integer = 3
Public Const Offset_Width As Integer = 7


Public Function Add_DCEs() As CopyRange
    Dim dce As New CopyRange
    
    With dce
        .startAddress = "A2"
        .SheetName = "DataImReadingIn"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\mjdexample.xlsx"
    End With
    
    Set Add_DCEs = dce
End Function

Public Function Add_Downlinks() As CopyRange
    Dim downlink As New CopyRange
    
    With downlink
        .startAddress = "D2"
        .SheetName = "DataImReadingIn"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\mjdexample.xlsx"
        .IsOffsetRange = True
    End With
    
    Set Add_Downlinks = downlink
End Function

Public Function Copy_To_Range() As CopyRange
    Dim copyToWorkbook As New CopyRange
    
    With copyToWorkbook
        .startAddress = "A2"
        .SheetName = "Sheet1"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\ForJames.xlsx"
    End With
    
    Set Copy_To_Range = copyToWorkbook
End Function
