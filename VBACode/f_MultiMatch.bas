Attribute VB_Name = "f_MultiMatch"
Function cMultiMatch(LookupValue As String, InitialRange As Range, ReturnHeader As String, ReturnRange As Range)

'Error Handler
On Error GoTo ErrorHandler1

'Obtains LookupRange Boundaries
LookupWorksheet = InitialRange.Worksheet.Name
LookupWorksheet_Excel = LookupWorksheet & "!"
ReturnWorksheet = ReturnRange.Worksheet.Name
ReturnWorksheet_Excel = ReturnWorksheet & "!"
LookupRange = InitialRange.Address
ReturnRangeAddress = ReturnRange.Address
StartAddress = Left(LookupRange, WorksheetFunction.Search(":", LookupRange) - 1)
StartRow = Range(StartAddress).Row
StartColumn = Range(StartAddress).Column
EndAddress = Right(LookupRange, Len(LookupRange) - Len(StartAddress) - 1)
EndRow = Range(EndAddress).Row
EndColumn = Range(EndAddress).Column
ReturnRangeAddress = ReturnRange.Address
ReturnStartAddress = Left(ReturnRangeAddress, WorksheetFunction.Search(":", ReturnRangeAddress) - 1)
ReturnStartColumn = Range(ReturnStartAddress).Column
ReturnStartRow = Range(ReturnStartAddress).Row

If LookupValue = "" Then
    cMultiMatch = "Key is Blank"
    Exit Function
End If

'Determines how many times the code needs to run
On Error Resume Next
    'Checks to see if the input is a number, if it is a number then search without asterisks as asterisks break a number countif for some reason.
    Checkifisnumber = LookupValue * 2
On Error GoTo ErrorHandler1
If Not IsError(Checkifisnumber) Then
    MaxResults = WorksheetFunction.CountIfs(Worksheets(LookupWorksheet).Range(LookupRange), LookupValue) - 1 'Subtracting 1 to fit array 'Asterisks might cause issues later with unwanted counts, don't think it will be an issue though
    Else
    MaxResults = WorksheetFunction.CountIfs(Worksheets(LookupWorksheet).Range(LookupRange), "*" & LookupValue & "*") - 1 'Subtracting 1 to fit array 'Asterisks might cause issues later with unwanted counts, don't think it will be an issue though
End If
If MaxResults < 0 Then
    cMultiMatch = ""
    Exit Function
End If

'Creates Array
Dim RowArray()
ReDim RowArray(0 To MaxResults)

'Finds Matching Instances
RangeAdjusttoTop = ReturnStartRow - 1
RangeAdjustFromLeft = ReturnStartColumn - 1
IncreaseRowBy = RangeAdjusttoTop
For i = 0 To MaxResults
ResultRow = Evaluate("Match(True, Index(IsNumber(Search(""" & LookupValue & """," & LookupWorksheet_Excel & LookupRange & ")), 0), 0)") + IncreaseRowBy 'old: ResultRow = WorksheetFunction.Match(True, WorksheetFunction.Index(WorksheetFunction.IsNumber(WorksheetFunction.Search(LookupValue, Worksheets(LookupWorksheet).Range(LookupRange))), 0), 0)) + IncreaseRowBy ResultRow = WorksheetFunction.Match(True, WorksheetFunction.Index(WorksheetFunction.IsNumber(WorksheetFunction.Search(LookupValue, Worksheets(LookupWorksheet).Range(LookupRange))), 0), 0) + IncreaseRowBy
RowArray(i) = Worksheets(ReturnWorksheet).Cells(ResultRow, Evaluate("Match(True, Index(IsNumber(Search(""" & ReturnHeader & """," & ReturnWorksheet_Excel & ReturnRangeAddress & ")), 0), 0)") + RangeAdjustFromLeft) 'old: Cells(ResultRow, WorksheetFunction.Match(ReturnHeader, ReturnRange, 0))
    LookupRange = Range(Cells(StartRow + ResultRow - RangeAdjusttoTop, StartColumn), Cells(EndRow, EndColumn)).Address 'make sure this functions as expected across worksheets
    IncreaseRowBy = ResultRow
    c = c + 1
Next i
'Rebuilds Array Into string
For i = 0 To c - 1
    RowOccurences = RowOccurences & RowArray(i) & ", "
Next i
'Creates Function Return
cMultiMatch = "Found the following (" & ReturnHeader & ") matches: " & RowOccurences
Exit Function

'Handles Errors
ErrorHandler1:
cMultiMatch = "Something went wrong." & Chr(10) & Err.Description & Error & Err.Number & Err.Source
End Function


