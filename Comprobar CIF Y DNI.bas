Attribute VB_Name = "CIFYDNI"

'*********************************
'Rutina de comprobación de DNI/CIF.
'*********************************
'Comprueba el DNI/CIF, y en caso de que este no lleve la letra, la devuelve
'si ya tiene la letra devuelve ""
Public Sub comprueba_DNI(DNICIF As String, objeto As Object)
Dim A As Integer, B As Integer, C As Integer, D As Integer, e As Integer
Dim nif As String, TpS As String * 1
nif = UCase(DNICIF)

If (Len(nif) = 8 And IsNumeric(nif)) Or (Len(nif) = 9 And Right(nif, 1) > "A") Then 'Es un carnet de identidad español
   
   A = Left(nif, 8) Mod 23
   TpS = Mid("TRWAGMYFPDXBNJZSQVHLCKET", A + 1)
   
   On Error Resume Next
   If TpS <> Right(nif, 1) Then
    If (Len(nif) = 8 And IsNumeric(nif)) Then objeto.Text = objeto.Text & TpS     'devolvemos la letra de control
    If (Len(nif) = 9 And Right(nif, 1) > "A") Then Right(objeto.Text, 1) = TpS
    MsgBox "La letra del control del DNI debiera ser la " & TpS, vbExclamation, "Comprobación de DNI/CIF"
   End If
   Err = 0
   On Error GoTo 0

Else  'Es una sociedad española
   If nif Like "ES-*" Then nif = Mid(nif, 4) 'Por si es intertacional de España
   If nif Like "ES*" Then nif = Mid(nif, 3)  'Por si es intertacional de España
   If Len([nif]) <> 9 Then
      MsgBox "El CIF parece ser incorrecto para una sociedad por no tener 9 dígitos (puede ser un CIF internacional)", vbInformation, "Comprobación de DNI/CIF"
      
   Else
      TpS = Left([nif], 1)
      If TpS <> "A" And TpS <> "B" And TpS <> "C" And TpS <> "D" And TpS <> "E" And _
         TpS <> "F" And TpS <> "G" And TpS <> "H" And TpS <> "N" And TpS <> "P" And TpS <> "Q" And TpS <> "S" Then
         MsgBox "El CIF parece ser incorrecto para una sociedad. La letra no es correcta", vbExclamation, "Comprobación de DNI/CIF"
       
      Else
         If Abs((Mid([nif], 2, 1))) > 4 Then A = 1 Else A = 0
         If Abs((Mid([nif], 4, 1))) > 4 Then B = 1 Else B = 0
         If Abs((Mid([nif], 6, 1))) > 4 Then C = 1 Else C = 0
         If Abs((Mid([nif], 8, 1))) > 4 Then D = 1 Else D = 0
         e = Abs(((Abs((Mid([nif], 2, 1))) + Abs((Mid([nif], 3, 1))) + _
            Abs((Mid([nif], 4, 1))) + Abs((Mid([nif], 5, 1))) + Abs((Mid([nif], 6, 1))) + _
            Abs((Mid([nif], 7, 1))) + Abs((Mid([nif], 8, 1))) + _
            Abs((Mid([nif], 2, 1))) + Abs((Mid([nif], 4, 1))) + Abs((Mid([nif], 6, 1))) + _
            Abs((Mid([nif], 8, 1))) + A + B + C + D) Mod 10) - 10)
         'Añadido por César al ver algún caso erróneo
Menos10:
         If e > 9 Then e = e - 10: GoTo Menos10
         'Añadido por César al ver algún caso erróneo
         If e <> Abs((Mid([nif], 9, 1))) Then
            MsgBox "El CIF parece ser incorrecto para una sociedad. El dígito final debiera ser un" & [e] & "", vbExclamation, "Comprobación de DNI/CIF"
            
         End If
      End If
   End If
End If
End Sub
