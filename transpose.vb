Public Sub Transpose()
    'Declare Variables
    Dim SourceTable As Table
    Dim RowCount As Long, ColumnCount As Long
    Dim TableRange As Range
    Dim i As Long, j As Long 'Loop Counters
    Dim RowDataAsArray() As String
    Dim NewTable As Table
    Dim SourceTableStyle As Style
    Dim TableAsArray() As String 'Will contain the table text in memory
    
    'If cursor is not in a table, exit macro
    If Not Selection.Information(wdWithInTable) Then
      MsgBox "Cursor should be in a table"
      Exit Sub
    End If
    
    Set SourceTable = Selection.Tables(1)
    RowCount = SourceTable.Rows.Count
    ColumnCount = SourceTable.Columns.Count
    Set SourceTableStyle = SourceTable.Style
 
   'Redefine array as a two dimensional array with exact row and column count
    ReDim TableAsArray(1 To RowCount, 1 To ColumnCount)
    
    For i = 1 To RowCount
      RowDataAsArray = Split(Expression:=SourceTable.Rows(i).Range.Text, _
                  Delimiter:=vbCr)
      For j = 1 To ColumnCount
        'Last item in RowDataAsArray is vbCr, thus j - 1 to ignore that
        TableAsArray(i, j) = RowDataAsArray(j - 1)
      Next j
    Next
    
    Set TableRange = SourceTable.Range
    TableRange.Collapse wdCollapseEnd
    SourceTable.Delete
    
    'Create a new table at the same position
    Set NewTable = TableRange.Tables.Add(TableRange, ColumnCount, RowCount)
    'Fill data in the new table
    For i = 1 To RowCount
      For j = 1 To ColumnCount
        NewTable.Rows(j).Cells(i).Range.Text = TableAsArray(i, j)
      Next
    Next
    'Apple Style to the new table
    NewTable.Style = SourceTableStyle
  End Sub
