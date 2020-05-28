Attribute VB_Name = "Módulo1"
Sub FiltroAvanzado()
Attribute FiltroAvanzado.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' FiltroAvanzado Macro
'
' Acceso directo: CTRL+f
'
    Range("A7:G23").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("A4:G5"), Unique:=False
End Sub
