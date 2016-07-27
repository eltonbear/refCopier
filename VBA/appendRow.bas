Attribute VB_Name = "appendRow"
Sub appendARow()
    Dim statusC As String
    Dim refC As String
    Dim copyC As String
    Dim typeC As String
    Dim depC As String
    Dim row As String
    Dim entireRow As String
    Dim prevAddedRowCell As String
    Dim lastRef As Integer
    Dim unlockedCell As String
    Dim lastAppendRowCell As String
    Dim lastAppendRow As Integer

    'set columns
    statusC = "A"
    refC = "B"
    copyC = "C"
    typeC = "D"
    depC = "E"
    'set cell locations
    prevAddedRowCell = "F2"
    lastAppendRowCell = "F3"
    'last added row
    prevAddedRow = CStr(Range(prevAddedRowCell).Value)
    'current row
    row = CStr(CInt(prevAddedRow) + 1)
    'maximum row to append to
    lastAppendRow = Range(lastAppendRowCell).Value 'integer
    
    If CInt(row) <= lastAppendRow Then
        'last refence number
        lastRef = Range(refC + prevAddedRow).Value
        'range for the whole row
        entireRow = statusC + row + ":" + depC + row
        'range for unlocked cells
        unlockedCell = copyC + row + ":" + typeC + row
        'lock sheet but allow vba to modify cells
        Call Worksheets(1).Protect(Password:="elton", UserInterfaceOnly:=True)
        'add formula on dep
        Range(depC + row).Formula = "=" + copyC + row
        'set fong color
        Range(statusC + row + ":" + refC + row).Font.Color = RGB(255, 255, 255)
        Range(depC + row).Font.Color = RGB(0, 0, 0)
        'set text alignment
        Range(entireRow).HorizontalAlignment = xlCenter
        'set row back groung color
        Range(entireRow).Interior.Color = RGB(146, 205, 220)
        'write status
        Range(statusC + row) = "appending"
        'write Ref
        Range(refC + row).Value = lastRef + 1
        'unlock for user to modify
        Range(unlockedCell).Locked = False
        'set border color and line style
        With Range(entireRow).Borders
            .LineStyle = xlContinuous
            .Color = RGB(178, 178, 178)
        End With
        'hide formula
        Range(depC + row).FormulaHidden = True
        'write current row back into a cell
        Range(prevAddedRowCell) = row
    Else
        'warning
        MsgBox "Exceed max Reference Number!", vbExclamation, "Warning"
    End If
    
End Sub

