Attribute VB_Name = "Module1"
Sub MergeCitiesData()
Attribute MergeCitiesData.VB_ProcData.VB_Invoke_Func = "x\n14"
    MergeCities.Show
End Sub

Sub LabelLocations()
Attribute LabelLocations.VB_ProcData.VB_Invoke_Func = "X\n14"
    For Each row In Selection.Rows
        ComputeLocation row
    Next row
End Sub

Private Sub ComputeLocation(row)
    Dim x As String
    x = row.Columns(3)
  
    row.Columns(1) = ""
    row.Columns(2) = ""
    
    'RDA, EDA, FDA, IRI
    Dim reserves As Variant
    reserves = Array(" RDA ", " EDA ", " FDA ", " IRI ")
    
    'VL, CY, DM, T
    Dim cities As Variant
    cities = Array(" VL ", " CY ", " DM ", " T ", " RGM ")
        
    If Find(x, reserves) = True Then
        'Do Nothing
    ElseIf Find(x, cities) = True Then
        row.Columns(2) = "1"
    Else
        row.Columns(1) = "1"
    End If
    
    
End Sub
 
Private Function Find(word As String, list As Variant) As Boolean
    For Each element In list
        If InStr(word, element) Then
            Find = True
            Exit Function
        End If
    Next element
    
    Find = False
End Function
