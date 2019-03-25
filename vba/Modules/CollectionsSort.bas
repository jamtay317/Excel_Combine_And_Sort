Attribute VB_Name = "CollectionsSort"
Public Sub SortCollection(col As Collection, psSortPropertyName As String, pbAscending As Boolean, Optional psKeyPropertyName As String)

Dim obj As Object
Dim i As Integer
Dim j As Integer
Dim iMinMaxIndex As Integer
Dim vMinMax As Variant
Dim vValue As Variant
Dim bSortCondition As Boolean
Dim bUseKey As Boolean
Dim sKey As String
    
    bUseKey = (psKeyPropertyName <> "")
    
    For i = 1 To col.Count - 1
        Set obj = col(i)
        vMinMax = CallByName(obj, psSortPropertyName, VbGet)
        iMinMaxIndex = i
        
        For j = i + 1 To col.Count
            Set obj = col(j)
            vValue = CallByName(obj, psSortPropertyName, VbGet)
            
            If (pbAscending) Then
                bSortCondition = (vValue < vMinMax)
            Else
                bSortCondition = (vValue > vMinMax)
            End If
            
            If (bSortCondition) Then
                vMinMax = vValue
                iMinMaxIndex = j
            End If
            
            Set obj = Nothing
        Next j
        
        If (iMinMaxIndex <> i) Then
            Set obj = col(iMinMaxIndex)
            
            col.Remove iMinMaxIndex
            If (bUseKey) Then
                sKey = CStr(CallByName(obj, psKeyPropertyName, VbGet))
                col.Add obj, sKey, i
            Else
                col.Add obj, , i
            End If
            
            Set obj = Nothing
        End If
        
        Set obj = Nothing
    Next i
        
End Sub
