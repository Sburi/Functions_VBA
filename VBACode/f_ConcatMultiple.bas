Attribute VB_Name = "f_ConcatMultiple"
Function cConcatMultiple(Ref As Range, Delimiter As String) As String

On Error GoTo ErrorHandler

Dim Cell As Range
Dim Result As String

For Each Cell In Ref
    If IsError(Cell.Value) Then
        'Result = Result & " " & Delimiter 'Remove comment if you want to see blank results
        'Do nothing assuming the above line is commented
    ElseIf Cell.Value = "" Then
        'Do nothing
    Else
        Result = Result & Cell.Value & Delimiter
    End If
Next Cell

If Result = "" Then
    cConcatMultiple = "No Results"
    Else
    cConcatMultiple = Left(Result, Len(Result) - 1)
End If

Exit Function

ErrorHandler:
cConcatMultiple = "Something went wrong. " & vbCrLf & vbError & vbCrLf & Error & vbCrLf
 
End Function
