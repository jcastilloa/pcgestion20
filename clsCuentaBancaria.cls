VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'****************************************
' (c) Carlos Augusto Barreira Díaz
' Fecha Creación: 03/06/97
'
' Finalidad:  Proporcionar una clase para validar números de Cuenta Bancario.
'
' Premisas y efectos: Su funcinamiento es totalmente autonomo.
'
' Funcionamiento: Hay dos formas de usar esta clase:
' 1.- Asignar a la Propiedad Cuenta (Que es la propiedad predeterminada)
' un valor de tipo string (es igual el formato ya que elimina todo lo que no
' sean numeros) y luego comprobar la propiedad Err.
' 2.- Usar la función Validar pasandole el nº de cuenta en un string y opcionalmente
' si se quiere pasar un boolean para dar formato al nº de cuenta (por defecto es true).
' Esta función devuelve el valor de la propiedad Err.
'
'**************************************

Option Explicit

Private Const strSep = "-"

Private mstrCuenta As String
Private mblnErr As Boolean
Private mblnDarFormato As Boolean
Private mintPesos(1 To 10) As Integer
Private strText1 As String, strBanco As String, strOficina As String, strCuenta As String, strDC1 As String, strDC2 As String
Private strDC1Bueno As String, strDC2Bueno As String

Public Property Get Err() As Boolean
  Err = mblnErr
End Property

Public Property Get DarFormato() As Boolean
'Attribute DarFormato.VB_Description = "Si es True (Valor por defecto) se da formato al nº de cuenta despues de validar correctamente el nº."'
  DarFormato = mblnDarFormato
End Property

Public Property Let DarFormato(blnDarFormato As Boolean)
  mblnDarFormato = blnDarFormato
End Property

Public Property Let Cuenta(ByVal vData As String)
  mstrCuenta = vData
  Validar_Todo
End Property

Public Property Get Cuenta() As String
'Attribute Cuenta.VB_UserMemId = 0
  Cuenta = mstrCuenta
End Property

Public Function Validar(strCuentaBanco As String, Optional blnDarFormato As Boolean = True) As Boolean
  mstrCuenta = strCuentaBanco
  mblnDarFormato = blnDarFormato
  Validar_Todo
  Validar = mblnErr
  strCuentaBanco = mstrCuenta
End Function

Private Sub Class_Initialize()
  mstrCuenta = ""
  mblnErr = True: mblnDarFormato = True
  mintPesos(1) = 6: mintPesos(2) = 3: mintPesos(3) = 7: mintPesos(4) = 9: mintPesos(5) = 10
  mintPesos(6) = 5: mintPesos(7) = 8: mintPesos(8) = 4: mintPesos(9) = 2: mintPesos(10) = 1
End Sub

Private Sub Validar_Todo()
  Dim i As Integer
  mblnErr = False
  strText1 = "" 'Eliminar todo menos los numeros
  For i = 1 To Len(mstrCuenta)
    If IsNumeric(Mid$(mstrCuenta, i, 1)) Then
      strText1 = strText1 + Mid$(mstrCuenta, i, 1)
    End If
  Next
  If Len(strText1) <> 20 Then
    mblnErr = True
   Else
    strBanco = Mid(strText1, 1, 4)
    strOficina = Mid(strText1, 5, 4)
    strDC1 = Mid(strText1, 9, 1)
    strDC2 = Mid(strText1, 10, 1)
    strCuenta = Mid(strText1, 11, 10)
    If Not Validar_cuenta(strBanco & strOficina, strDC1, False, strDC1Bueno) Then
     mblnErr = True
    ElseIf Not Validar_cuenta(strCuenta, strDC2, True, strDC2Bueno) Then
     mblnErr = True
    ElseIf mblnDarFormato Then
      mstrCuenta = strBanco & strSep & strOficina & strSep & strDC1 & strDC2 & strSep & strCuenta
    End If
  End If
End Sub

Private Function Validar_cuenta(ByVal num As String, ByVal cc As String, ByVal Cuenta As Boolean, Optional strDcCorrecto As String) As Boolean
  Dim contpesos As Integer, contnum As Integer, suma As Long, resto As Integer, ccc As Integer

  If Cuenta Then
    contpesos = 10 'se paso un nº de cuenta
   Else
    contpesos = 8 ' se paso el nº de banco y oficina
  End If

  For contnum = 1 To Len(num)
    suma = suma + (mintPesos(contpesos) * CLng(Mid(num, contnum, 1)))
    contpesos = contpesos - 1
  Next
  resto = suma Mod 11
  ccc = 11 - resto
  If ccc = 10 Then ccc = 1
  If ccc = 11 Then ccc = 0
  If ccc = CInt(cc) Then
    Validar_cuenta = True
   Else
    Validar_cuenta = False
  End If
  If Not IsMissing(strDcCorrecto) Then
    strDcCorrecto = CStr(ccc)
  End If
End Function

