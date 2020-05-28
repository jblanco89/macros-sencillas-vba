VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CmdAddsheet_Click()


Sheets.Add.Name = ActiveCell.Value

End Sub

Private Sub cmdeliminarhojas_Click()
Dim boton As Integer
boton = MsgBox("¿desea eliminar esta hoja?", vbCritical + vbYesNo + vbDefaultButton2, "Advertencia")
If boton = vbYes Then
    Application.DisplayAlerts = False
    Worksheets(ActiveCell.Value).Delete
    Application.DisplayAlerts = True
    ActiveCell.Delete
End If

End Sub

