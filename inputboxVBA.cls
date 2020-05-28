VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdadd_Click()
Dim Fecha As Date
Dim Planta As String
Dim Mensaje As String
Dim ProductoA As Double
Dim ProductoB As Double
Dim ProductoC As Double
Dim Total As Double
Dim btn As Integer
Dim Fila As Long

        Fecha = InputBox("Ingrese Fecha", "Fecha", Date)
        Planta = InputBox("Ingrese nombre de Planta", "Planta", "Alicante")
        ProductoA = InputBox("Ingrese producto A.", "Producto A.")
        ProductoB = InputBox("Ingrese producto B.", "Producto B.")
        ProductoC = InputBox("Ingrese producto C.", "Producto C.")
        Total = ProductoA + ProductoB + ProductoC
        
        Mensaje = "Verifica los datos, ¿Son correctos?" & vbCrLf
        Mensaje = Mensaje & "Fecha: " & Fecha & vbCrLf
        Mensaje = Mensaje & "Planta: " & Planta & vbCrLf
        Mensaje = Mensaje & "Producto A: " & ProductoA & vbCrLf
        Mensaje = Mensaje & "Producto B: " & ProductoB & vbCrLf
        Mensaje = Mensaje & "Producto C: " & ProductoC & vbCrLf
        Mensaje = Mensaje & "Total: " & ProductoA + ProductoB + ProductoC & vbCrLf
        
        btn = MsgBox(Mensaje, vbInformation + vbYesNoCancel, "Guardar")
        
        If btn = vbYes Then
            Fila = ActiveSheet.Cells(1048576, "A").End(xlUp).Row + 1
            Cells(Fila, "A").Value = Fecha
            Cells(Fila, "B").Value = Planta
            Cells(Fila, "C").Value = ProductoA
            Cells(Fila, "D").Value = ProductoB
            Cells(Fila, "E").Value = ProductoC
            Cells(Fila, "F").Value = Total
        End If
        
        
        End Sub

Private Sub cmddelete_Click()
Dim btn As Integer
btn = MsgBox("¿desea eliminar esta fila?", vbCritical + vbYesNo + vbDefaultButton2, "Advertencia")
If btn = vbYes Then
    ActiveCell.EntireRow.Delete
End If
End Sub
