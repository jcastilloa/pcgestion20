VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrystalFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Modulo      : clsCrystalFormula
' Fecha/Hora  : 19/08/2004 12:28
' Autor       : JCASTILLO
' Propósito   : Armar una formula para crystal report, una idea similar a clsSmartSQL pero
'               mas sencillo
'---------------------------------------------------------------------------------------
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarFormula As String 'copia local

Public Sub Borrar()

mvarFormula = ""

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : AñadeCampo
' Fecha/Hora  : 19/08/2004 12:29
' Autor       : JCASTILLO
' Propósito   : Añade un campo a la formula de crystal report
'
'---------------------------------------------------------------------------------------
Public Sub AñadeCampo(Tabla As String, NombreCampo As String, Valor As String, Operador As String, Clausula_AND_OR As String, Optional Es_Fecha As Boolean, Optional Es_Booleano As Boolean, Optional es_char As Boolean)
Dim tmpcampo As String

   On Error GoTo AñadeCampo_Error

If Trim(mvarFormula) <> "" Then mvarFormula = mvarFormula & " " & Clausula_AND_OR

If Es_Fecha Then
    tmpcampo = "date(" & Right(Valor, 4) & "," & Format(Valor, "mm") & "," & Format(Valor, "dd") & ")"
    mvarFormula = mvarFormula & " " & "({" & Tabla & "." & NombreCampo & "} " & Operador & " " & tmpcampo & ")"
ElseIf Es_Booleano Then

    If Valor = "0" Then
        mvarFormula = mvarFormula & " " & "({" & Tabla & "." & NombreCampo & "} " & Operador & " " & "False" & ")"
    Else
        mvarFormula = mvarFormula & " " & "({" & Tabla & "." & NombreCampo & "} " & Operador & " " & "True" & ")"
    End If
    
Else
    
    If es_char Then
    
        If UCase(Operador) = "LIKE" Then
    
            Valor = "*" & Valor & "*"  'si el operador es like colocar los *
    
        End If
        
        Valor = """" & Valor & """"
    
    End If
    
    mvarFormula = mvarFormula & " " & "({" & Tabla & "." & NombreCampo & "} " & Operador & " " & Valor & ")"
End If

   On Error GoTo 0
   Exit Sub

AñadeCampo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento AñadeCampo de Módulo de clase clsCrystalFormula"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : AñadeClausulaLibre
' Fecha/Hora  : 19/08/2004 12:30
' Autor       : JCASTILLO
' Propósito   : Añadir una clausula compleja libremente
'
'---------------------------------------------------------------------------------------
Public Sub AñadeClausulaLibre(Clausula As String, Clausula_AND_OR As String)

   On Error GoTo AñadeClausulaLibre_Error

If Trim(mvarFormula) <> "" Then mvarFormula = mvarFormula & " " & Clausula_AND_OR
mvarFormula = mvarFormula & " " & "(" & Clausula & ")"

   On Error GoTo 0
   Exit Sub

AñadeClausulaLibre_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento AñadeClausulaLibre de Módulo de clase clsCrystalFormula"

End Sub


Public Property Get formula() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Formula
    formula = mvarFormula
End Property



