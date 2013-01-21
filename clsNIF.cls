VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNIF"
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
' Finalidad:  Proporcionar una clase para validar y calcular letra del NIF.
'
' Premisas y efectos: Su funcinamiento es totalmente autonomo.
'
' Funcionamiento: La clase puede validar y calcular la letra del nif para
' ello asume que si la ultima letra del string que se le pasa es una letra valida
' lo que se quiere es validar el nif y si no lo que se quiere es calcular el NIF
' el resto de caracteres solo son significativos si son numericos. Así pues el
' funcionamiento es de dos formas:
' 1.- Asignar a la Propiedad NIF (Que es la propiedad predeterminada)
' un valor de tipo string (es igual el formato ya que elimina todo lo que no
' sean numeros) y luego comprobar la propiedad Err.
' 2.- Usar la función Validar pasandole el nº de NIF en un string y opcionalmente
' si se quiere dar formato al nº de cuenta (por defecto es true). Esta función
'  devuelve el valor de la propiedad Err.
' En ambos casos la propiedad NIF contiene el NIF resultado segun los calculos
' o formato indicados.
'**************************************

Option Explicit

Private Const strLetras = "TRWAGMYFPDXBNJZSQVHLCKE"
Private Const strSep = "-"

Private mstrNIF As String 'copia local
Private mstrLetra As String, mstrLetraBuena As String
Private lngNIF As Long
Private blnErr As Boolean, blnFormat As Boolean
Public Property Let DarFormato(blnFor As Boolean)
  blnFormat = blnFor
End Property

Public Property Get DarFormato() As Boolean
  DarFormato = blnFormat
End Property

Public Property Get Err() As Boolean
  Err = blnErr
End Property

Public Property Get Letra() As String
  Letra = mstrLetra
End Property

Public Property Let nif(ByVal vData As String)
  mstrNIF = vData
  Validar_NIF
End Property

Public Property Get nif() As String

  nif = mstrNIF
End Property

Public Function Validar(nif As String, Optional Format As Boolean = True) As Boolean
  mstrNIF = nif
  blnFormat = Format
  Validar_NIF
  Validar = Err
  nif = mstrNIF
End Function

Private Sub Validar_NIF()
  Dim strN As String
  
  blnErr = False
  strN = Dejar_Numeros(mstrNIF)
  If strN = "" Then
    blnErr = True
    mstrLetraBuena = "" ': mstrNIF = ""
   Else
    lngNIF = CLng(strN)
    mstrLetra = UCase(Mid(mstrNIF, Len(mstrNIF), 1))
    mstrLetraBuena = CalcularLetra
    If InStr(strLetras, mstrLetra) Then
      'se desea validar
      If mstrLetra <> mstrLetraBuena Then
        blnErr = True ': mstrNIF = ""
      End If
     Else
      mstrNIF = mstrNIF + mstrLetraBuena
    End If
    'si no se produjo error
    If (Not blnErr) And blnFormat Then
      mstrNIF = Format(lngNIF, "00,000,000") & strSep & mstrLetraBuena
    End If
  End If
End Sub

Private Function CalcularLetra() As String
  CalcularLetra = Mid(strLetras, 1 + (lngNIF Mod 23), 1)
End Function

Private Function Dejar_Numeros(strN As String) As String
  Dim i As Integer
  For i = 1 To Len(strN)
    If IsNumeric(Mid$(strN, i, 1)) Then
      Dejar_Numeros = Dejar_Numeros + Mid$(strN, i, 1)
    End If
  Next
End Function

Private Sub Class_Initialize()
  blnFormat = True
End Sub


