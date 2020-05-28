VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 3 And Target.Row >= 2 Then
        If Target.Offset(0, -1) = "Si" And Target >= Range("H3") Then
        Target.Offset(0, 1) = Range("H1")
        Else
        Target.Offset(0, 1) = Range("H2")
        End If
    Target.Offset(0, 2) = Target * Target.Offset(0, 1)
    Target.Offset(1, -2).Select
    End If
End Sub

