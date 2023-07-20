Attribute VB_Name = "Module1"
Sub Cut_delete_row()
Attribute Cut_delete_row.VB_ProcData.VB_Invoke_Func = "G\n14"
    Dim selectedCell As Range
    Dim ws As Worksheet

    Set ws = ActiveSheet ' Change this if you want to work with a specific worksheet
    
    ' Check if a cell is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell before running this macro.", vbExclamation
        Exit Sub
    End If
    
    ' Check if only one cell is selected
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select only one cell before running this macro.", vbExclamation
        Exit Sub
    End If
    
    ' Store the selected cell
    Set selectedCell = Selection
    
    ' Check if the selected cell is in row 2 or below
    If selectedCell.Row <= 1 Then
        MsgBox "Please select a cell in row 2 or below.", vbExclamation
        Exit Sub
    End If
    
    ' Store the value of the selected cell
    Dim cellValue As Variant
    cellValue = selectedCell.Value
    
    ' Select the cell 2 rows below the selected cell (cell where we want to paste)
    Dim pasteCell As Range
    Set pasteCell = ws.Cells(selectedCell.Row + 2, selectedCell.Column)
    
    ' Paste the value in the cell 2 rows below the selected cell
    pasteCell.Value = cellValue
    
    ' Store the rows to delete (row1 and row3)
    Dim rowToDelete1 As Range
    Dim rowToDelete2 As Range
    Dim rowToDelete3 As Range
    Set rowToDelete1 = ws.Rows(selectedCell.Row)
    Set rowToDelete2 = ws.Rows(selectedCell.Row + 1)
    Set rowToDelete3 = ws.Rows(selectedCell.Row - 1)
    
    ' Delete the rows (row1 and row3)
    Application.DisplayAlerts = False ' To avoid the delete confirmation message
    rowToDelete1.Delete Shift:=xlUp
    rowToDelete2.Delete Shift:=xlUp
    rowToDelete3.Delete Shift:=xlUp
    Application.DisplayAlerts = True
    
    
    ' Clear the clipboard
    Application.CutCopyMode = False

    ' Find the last row in the column of the selected cell
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, pasteCell.Column).End(xlUp).Row

    ' Copy the value down until the next cell already has data in it
    Dim copyRange As Range
    Set copyRange = ws.Range(pasteCell, ws.Cells(lastRow, pasteCell.Column))

    Dim nextCell As Range
    For Each nextCell In copyRange
        If nextCell.Offset(1, 0).Value <> "" Then
            Exit For
        End If
        nextCell.Value = cellValue
    Next nextCell
End Sub


Sub SelectCellAboveBlankInColumnA()
Attribute SelectCellAboveBlankInColumnA.VB_ProcData.VB_Invoke_Func = "F\n14"
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    
    ' Set the worksheet where you want to search for the blank cell
    Set ws = ActiveSheet
    ' Find the last row with data in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the cells in Column A from top to bottom
    For Each cell In ws.Range("A2:A" & lastRow) ' Start from A2 to avoid selecting the header row
        ' Check if the cell is blank
        If cell.Value = "" Then
            ' Select the cell above the blank cell in Column A
            ws.Activate
            ws.Cells(cell.Row + 1, "A").Select
            Exit Sub ' Exit the loop once a blank cell is found
        End If
    Next cell
    
    
End Sub
Sub RunMacrosInLoop()
    Dim i As Integer
    Dim numberOfLoops As Integer
    
    numberOfLoops =  ' Change this value to the desired number of loops
    
    For i = 1 To numberOfLoops
        ' Call the function "SelectCellAboveBlankInColumnA()" directly and execute the code within it
        Call SelectCellAboveBlankInColumnA
        
        ' Call the subroutine "Cut_delete_row()" directly and execute the code within it
        Call Cut_delete_row
    Next i
End Sub
Sub RunMacrosInLoop_auto()
    Dim i As Integer
    Dim cellAboveBlank As Range
    Dim loopLimit As Integer
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    
    ' Set the worksheet where you want to search for the blank cell
    Set ws = ActiveSheet
    ' Find the last row with data in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Set the maximum number of loops to a large value to avoid infinite loops
    loopLimit = 36
    
    For i = 1 To loopLimit
            
        ' Call the function "SelectCellAboveBlankInColumnA()" directly and execute the code within it
        Call SelectCellAboveBlankInColumnA
        
        ' Call the subroutine "Cut_delete_row()" directly and execute the code within it
        Call Cut_delete_row
        

    
  
    

    Next i
End Sub

        
      





