Attribute VB_Name = "M_ConcatNonDuplicates"
Function cConcatNonDuplicates(Delimiter As String, ParamArray CellSelectionArray() As Variant)

On Error GoTo ErrHandler

'Builds a unique array of entries
i = 0
Dim UnqArray() As Variant
ReDim UnqArray(0 To UBound(CellSelectionArray))
For Each Item In CellSelectionArray
    If Trim(Item) <> "" Then
        If IsInArray = Not IsError(Application.Match(Trim(Item), UnqArray, 0)) = True Then
            UnqArray(i) = Trim(Item)
            i = i + 1
        End If
    End If
Next Item

'If 0 entries, show blank. if 1 entry, show item alone with no delimiter
If i = 0 Then
    cConcatNonDuplicates = ""
    Exit Function
ElseIf i = 1 Then
    cConcatNonDuplicates = UnqArray(0)
    Exit Function
End If

'Redims the unique array to the actual items entered
    ReDim Preserve UnqArray(0 To i - 1)

'Builds string output based on the unique array
For Each Item In UnqArray
    cConcatNonDuplicates = Item & Delimiter & cConcatNonDuplicates
Next Item

Exit Function

ErrHandler:
cConcatNonDuplicates = Error & vbCrLf & Err

End Function
