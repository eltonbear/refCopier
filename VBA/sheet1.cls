VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim lastExistingRefRow As String
    lastExistingRefRow = CStr(Range("U" + CStr(Range("V1").Value)).Value + 1)
    
    If Not Intersect(Target, Range("E2:E" + lastExistingRefRow)) Is Nothing And Target.HasFormula = False Then '' add comment
        If Target.Locked = True Then
            Dim anwer As Integer
            answer = MsgBox("Are you sure you want to modify dependon?", vbYesNo + vbQuestion, "Warning")
            If answer = vbYes Then
                Target.Locked = False
                MsgBox ("Cell " + Target.Address(RowAbsolute:=False, ColumnAbsolute:=False) + " has been unlocked!")
            End If
        Else
            Target.Locked = True
            MsgBox ("Cell " + Target.Address(RowAbsolute:=False, ColumnAbsolute:=False) + " has been locked!")
        End If
    End If
    Cancel = True
End Sub
