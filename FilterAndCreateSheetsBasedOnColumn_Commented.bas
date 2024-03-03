' Subroutine to filter data based on a specified column and create new sheets for each unique value found in that column.
' @param columnName The name of the column to filter on. This is case-insensitive.
Sub FilterAndCreateSheetsBasedOnColumn(columnName As String)
    ' Declare variables for working with worksheets, rows, columns, and values.
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
    
    ' Set the source worksheet to the currently active sheet.
    Set wsSource = ActiveSheet
    
    ' Determine the last row and column with data to define the search range.
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Apply an autofilter to the top row, preparing for data filtering.
    wsSource.Rows(1).AutoFilter
    
    ' Attempt to find the specified column by name.
    colIndex = 0
    For i = 1 To lastCol
        If LCase(wsSource.Cells(1, i).Value) = LCase(columnName) Then
            colIndex = i
            Exit For
        End If
    Next i
    
    ' If the column wasn't found, alert the user and exit the subroutine.
    If colIndex = 0 Then
        MsgBox "Column '" & columnName & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Collect unique values from the specified column.
    Set uniqueValues = New Collection
    On Error Resume Next ' Ignore errors to handle duplicate items in the collection.
    For Each cell In wsSource.Range(wsSource.Cells(2, colIndex), wsSource.Cells(lastRow, colIndex))
        uniqueValues.Add cell.Value, CStr(cell.Value)
    Next cell
    On Error GoTo 0 ' Resume normal error handling.
    
    ' Create a new sheet for each unique value and copy relevant data to it.
    For Each value In uniqueValues
        sheetNameBase = CleanSheetName(CStr(value))
        sheetName = sheetNameBase
        counter = 1
        
        ' Ensure the sheet name is unique by appending a counter if necessary.
        While SheetExists(sheetName)
            sheetName = Left(sheetNameBase, 31 - Len(CStr(counter)) - 1) & "_" & counter
            counter = counter + 1
        Wend
        
        ' Create the new sheet and name it.
        Set wsDest = Sheets.Add(After:=Sheets(Sheets.Count))
        wsDest.Name = sheetName
        
        ' Copy filtered data to the new sheet.
        wsSource.AutoFilterMode = False
        wsSource.Rows(1).AutoFilter Field:=colIndex, Criteria1:=value
        wsSource.AutoFilter.Range.Copy Destination:=wsDest.Range("A1")
        
        ' Reset the autofilter.
        wsSource.AutoFilterMode = False
    Next value
    
    ' Notify the user upon completion.
    MsgBox "Sheets created for each unique '" & columnName & "' value's up to 31 characters.", vbInformation
End Sub

' Function to clean and validate sheet names according to Excel's constraints.
' Removes invalid characters and ensures the name does not exceed 31 characters.
' @param name The original sheet name which may contain invalid characters.
' @return A cleaned and truncated (if necessary) version of the sheet name.
Function CleanSheetName(name As String) As String
    Dim cleanedName As String
    
    ' Sequentially replace each set of invalid characters with an empty string.
    cleanedName = Replace(name, ":", "")
    cleanedName = Replace(cleanedName, "\", "")
    cleanedName = Replace(cleanedName, "/", "")
    cleanedName = Replace(cleanedName, "?", "")
    cleanedName = Replace(cleanedName, "*", "")
    cleanedName = Replace(cleanedName, "[", "")
    cleanedName = Replace(cleanedName, "]", "")
    
    ' Truncate the name to the maximum length allowed for Excel sheet names.
    If Len(cleanedName) > 31 Then
        cleanedName = Right(cleanedName, 31)
    End If
    
    CleanSheetName = cleanedName
End Function

' Function to check if a sheet with the specified name already exists in the workbook.
' @param sheetName The name of the sheet to check for.
' @return True if the sheet exists, False otherwise.
Function SheetExists(sheetName As String) As Boolean
    ' Check for the existence of the sheet and handle errors silently.
    SheetExists = Not Worksheets(sheetName) Is Nothing
End Function
