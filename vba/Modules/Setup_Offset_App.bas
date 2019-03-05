Attribute VB_Name = "Setup_Offset_App"
Option Explicit

'The column that apears first will be the column that sorted by
Public Const DownLink_Columns As String = "A,B,H"
Public Const DCE_Columns As String = "A,D,J"

'the first column is the date column
Public Const DownLink_CopyTo_Columns As String = "J,L,M"
Public Const DCE_CopyTo_Columns As String = "B,E,F"

'the total amount of columns that are coppied
Public Const Copy_Width As Integer = 3

'the amount of columns that will be offset from downlinks and dces
Public Const Offset_Width As Integer = 7

'this is the range of the dces that you're copping from
'    With dce
'        .startAddress = this is the first cell that you would like to copy
'        .SheetName = the name of the sheet that you would like to copy from
'        .WorkbookAddress = the physical address where the workbook located
'    End With
Public Function Add_DCEs() As CopyRange
    Dim dce As New CopyRange
    
    With dce
        .startAddress = "A2"
        .SheetName = "Dces"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\newitems.xlsx"
    End With
    
    Set Add_DCEs = dce
End Function

'this is the range of the downlinks that you're copping from
'    With DownLink
'        .startAddress = this is the first cell that you would like to copy
'        .SheetName = the name of the sheet that you would like to copy from
'        .WorkbookAddress = the physical address where the workbook located
'    End With
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

'where you would like to copy the workbook to
'    With copyToWorkbook
'        .startAddress = the first cell that you would like populated
'        .SheetName = the name of the sheet that you would like to copy to
'        .WorkbookAddress = the physical address where the workbook located
'    End With
Public Function Copy_To_Range() As CopyRange
    Dim copyToWorkbook As New CopyRange
    
    With copyToWorkbook
        .startAddress = "A3"
        .SheetName = "Sheet1"
        .WorkbookAddress = "C:\Users\James\Desktop\Excel_Combine_And_Sort\Examples\ForJames.xlsx"
    End With
    
    Set Copy_To_Range = copyToWorkbook
End Function
