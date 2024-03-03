Sub FilterAndCreateSheetsBasedOnColumn(columnName As String)
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long
    Dim colIndex As Long
    Dim uniqueValues As Collection
    Dim value As Variant
    Dim cell As Range
    Dim sheetName As String
    Dim sheetNameBase As String
    Dim counter As Integer
    
    Set wsSource = ActiveSheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Apply filter to the top row
    wsSource.Rows(1).AutoFilter
    
    ' Find the specified column
    colIndex = 0
    For i = 1 To lastCol
        If LCase(wsSource.Cells(1, i).Value) = LCase(columnName) Then
            colIndex = i
            Exit For
        End If
    Next i
    
    If colIndex = 0 Then
        MsgBox "Column '" & columnName & "' not found.", vbExclamation
        Exit Sub
    End If
    
    Set uniqueValues = New Collection
    On Error Resume Next
    For Each cell In wsSource.Range(wsSource.Cells(2, colIndex), wsSource.Cells(lastRow, colIndex))
        uniqueValues.Add cell.Value, CStr(cell.Value)
    Next cell
    On Error GoTo 0
    
    For Each value In uniqueValues
        sheetNameBase = CleanSheetName(CStr(value))
        sheetName = sheetNameBase
        counter = 1
        
        While SheetExists(sheetName)
            sheetName = Left(sheetNameBase, 31 - Len(CStr(counter)) - 1) & "_" & counter
            counter = counter + 1
        Wend
        
        Set wsDest = Sheets.Add(After:=Sheets(Sheets.Count))
        wsDest.Name = sheetName
        
        wsSource.AutoFilterMode = False
        wsSource.Rows(1).AutoFilter Field:=colIndex, Criteria1:=value
        wsSource.AutoFilter.Range.Copy Destination:=wsDest.Range("A1")
        
        wsSource.AutoFilterMode = False
    Next value
    
    MsgBox "Sheets created for each unique '" & columnName & "' value's up to 31 characters.", vbInformation
End Sub

Function CleanSheetName(name As String) As String
    ' Ensure the name is valid for an Excel sheet and does not exceed 31 characters
    CleanSheetName = Replace(name, ":", "")
    CleanSheetName = Replace(CleanSheetName, "\", "")
    CleanSheetName = Replace(CleanSheetName, "/", "")
    CleanSheetName = Replace(CleanSheetName, "?", "")
    CleanSheetName = Replace(CleanSheetName, "*", "")
    CleanSheetName = Replace(CleanSheetName, "[", "")
    CleanSheetName = Replace(CleanSheetName, "]", "")
    If Len(CleanSheetName) > 31 Then
        CleanSheetName = Right(CleanSheetName, 31)
    End If
End Function

Function SheetExists(sheetName As String) As Boolean
    ' Check if a sheet with the given name exists
    SheetExists = Not Worksheets(sheetName) Is Nothing
End Function
