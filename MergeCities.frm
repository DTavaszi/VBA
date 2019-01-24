VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MergeCities 
   Caption         =   "Merge data"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "MergeCities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MergeCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub CitiesListBox_Click()
    SubmitBtn.Enabled = True
End Sub

Private Function SelectedValue() As String
    For i = 0 To CitiesListBox.ListCount - 1
        If CitiesListBox.Selected(i) = True Then
            SelectedValue = CitiesListBox.list(i)
            Exit Function
        End If
    Next i
End Function

Private Sub Merge(value As String)
    'Set city name to selected value
    Selection.Rows(1).Columns(1) = value
    
    Dim FirstRow As Boolean
        FirstRow = True
        
    For FRCol = 2 To Selection.Rows(1).Columns.Count
        For Each R In Selection.Rows
            Dim FirstCol As Boolean
            FirstCol = True
            
            If FirstRow Then
                FirstRow = False
            Else
                Selection.Rows(1).Columns(FRCol) = Selection.Rows(1).Columns(FRCol) + R.Columns(FRCol)
            End If
            
            FirstCol = True
        Next R
        FirstRow = True
    Next FRCol
    
    Dim rUnselect As Range
    Dim rCell As Range
    Dim rNew As Range
    
    Set rUnselect = Selection.Rows(1)
    
    For Each rCell In Selection
        If Intersect(rUnselect, rCell) Is Nothing Then
            If rNew Is Nothing Then
                Set rNew = rCell
            Else
                Set rNew = Union(rNew, rCell)
            End If
        End If
    Next
    
    rNew.Delete
    'Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1, Selection.Columns.Count).Select
    'Selection.Remove
    
End Sub

Private Sub SubmitBtn_Click()
    Merge (SelectedValue())
    Unload Me
End Sub

Private Sub UserForm_Click()
    
End Sub

Private Sub UserForm_Initialize()
    For Each R In Selection.Rows
        With CitiesListBox
            .AddItem R.Columns(1)
        End With
    Next
End Sub
