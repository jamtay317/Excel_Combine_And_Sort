Attribute VB_Name = "Setup_Offset_App"
Option Explicit

'The column that apears first will be the column that sorted by
Public Const DownLink_Columns As String = "A,B,H"
Public Const DCE_Columns As String = "A,D,J"

'the first column is the date column
Public Const DownLink_CopyTo_Columns As String = "J,L,M"
Public Const DCE_CopyTo_Columns As String = "B,E,F"

Public Const Copy_Width As Integer = 3
Public Const Offset_Width As Integer = 7


Public Function Add_DCEs() As CopyRange
    Dim dce As New CopyRange
    
    With dce
        .startAddress = "A2"
        .SheetName = "Dces"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\newitems.xlsx"
    End With
    
    Set Add_DCEs = dce
End Function

Public Function Add_Downlinks() As CopyRange
    Dim DownLink As New CopyRange
    
    With DownLink
        .startAddress = "D2"
        .SheetName = "downlinks"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\newitems.xlsx"
        .IsOffsetRange = True
    End With
    
    Set Add_Downlinks = DownLink
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
