Attribute VB_Name = "undoAppend"
Sub undoRow()
    Dim statusC As String
    Dim refC As String
    Dim copyC As String
    Dim typeC As String
    Dim depC As String
    Dim row As String
    Dim entireRow As String
    Dim prevAddedRowCell As String
    Dim lockedCell As String
    Dim firstRowCell As String
    Dim firstRow As Integer
    
    'set columns
    statusC = "A"
    refC = "B"
    copyC = "C"
    typeC = "D"
    depC = "E"
    'set cell locations
    firstRowCell = "F1"
    prevAddedRowCell = "F2"
 
    'last added row
    prevAddedRow = CStr(Range(prevAddedRowCell).Value)
    'current row to delete is prev added row
    row = prevAddedRow
    'min row to delete to
    firstRow = Range(firstRowCell).Value 'integer
    If CInt(row) > firstRow Then
        'range for the whole row
        entireRow = statusC + row + ":" + depC + row
        'range for locked cells
        lockedCell = copyC + row + ":" + typeC + row
        'lock sheet but allow vba to modify cells
        Call Worksheets(1).Protect(Password:="elton", UserInterfaceOnly:=True)
        'unlock for user to modify
        Range(lockedCell).Locked = True
        'set border color and line style
        Range(entireRow).Borders.Color = RGB(217, 217, 217)
        'set row back groung color white
        Range(entireRow).Interior.Color = RGB(255, 255, 255)
        'set fong color white
        Range(depC + row).Font.Color = RGB(255, 255, 255)
        'write 'None' to dep
        Range(depC + row) = "None"
        'write status Null
        Range(statusC + row) = Null
        'write Ref Null
        Range(refC + row) = Null
        'write copy Null
        Range(copyC + row) = Null
        'write prev row back into a cell
        Range(prevAddedRowCell) = CStr(CInt(row) - 1)
    Else
        'warning
        MsgBox "Cannot reduce rows!", vbExclamation, "Warning"
    End If
    
End Sub

