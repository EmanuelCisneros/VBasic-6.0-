' Visual-Basic-6.0--Program-of-Methodology-Graphic

Dim COLOR As Long

Private Sub BOTONBORRAR_Click()
'limpiar el area

DIB.Cls
End Sub

Private Sub BOTONSALIR_Click()
If MsgBox("¿SALIR AHORA?", vbYesNo, "graficador") = vbYes Then
End
End If

End Sub

Private Sub COLORES_CLICK(INDEX As Integer)

Select Case INDEX
Case 0
COLOR = RGB(0, 0, 0) 'NEGRO
Case 1
COLOR = RGB(0, 0, 255) 'AZUL
Case 2
COLOR = RGB(255, 0, 0) 'ROJO
Case 3
COLOR = RGB(255, 255, 0) 'AMARILLO
Case 4
COLOR = RGB(0, 255, 0) ' VERDE
End Select
End Sub


Private Sub DIB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And opciones(0).Value = True Then
DIB.FillColor = QBColor(Rnd * 15) 'color de relleno aleatorio
DIB.FillStyle = Int(Rnd * 8) 'estilo de relleno aleatorio
DIB.Circle (x, y), 1000, COLOR 'dibuja el circulo
End If
If Button = 1 And opciones(1).Value = True Then
DIB.Line -(x, y), COLOR
End If
If Button = 1 And opciones(3).Value = True Then
DIB.Line -(x, y), COLOR, B
End If
If Button = 1 And opciones(4).Value = True Then
DIB.Line -(x, y), COLOR, BF
End If
End Sub

Private Sub DIB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And opciones(2).Value = True Then
DIB.DrawWidth = 5 'aumenta el tamaño del puento
DIB.PSet (x, y), COLOR
End If
If Button = 2 Then
DIB.DrawWidth = 20
DIB.PSet (x, y), RGB(255, 255, 255)
DIB.DrawWidth = 5
End If
End Sub

