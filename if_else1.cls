VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Primeramente nos aseguramos de usar el objeto Worksheet y seleccionar el evento "Change" en las declaraciones
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 1 And Target.Row >= 2 Then
        If IsDate(Target.Value) = True Then
            Target.Offset(0, 1) = StrConv(Format(Target, "dddd"), vbProperCase) 'StrConv es para colocar la primera letra mayúscula
            If Weekday(Target) = 7 Or Weekday(Target) = 1 Then
            Target.Interior.Color = RGB(249, 139, 127)
            Else
            Target.Interior.Color = xlNone
            End If
        ElseIf Target <> "" Then 'en caso de que la celda esté vacía
        MsgBox "Asegúrese de colocar una fecha, por favor corregirlo", vbCritical, "Error de valor"
        Target.Select
        End If
    End If

End Sub

