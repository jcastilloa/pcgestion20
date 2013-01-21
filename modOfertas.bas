Attribute VB_Name = "modOfertas"
'---------------------------------------------------------------------------------------
' Módulo       : modOfertas
' Fecha/Hora : 21/06/2004 10:50
' Autor         : JCastillo
' Propósito   :  Rutinas para manejar las ofertas del programa
'---------------------------------------------------------------------------------------
Option Explicit


'guarda el tipo de oferta a aplicar actualmente
Public OfertaActual As Byte
'descripcion de la oferta
Public OfertaDSC As String
'dcto por defecto de la oferta
Public OfertaDcto As Byte
'importe por defecto
Public OfertaImp As Double



'---------------------------------------------------------------------------------------
' Procedimiento : Comprueba_Ofertas
' Fecha/Hora     : 21/06/2004 10:47
' Autor             : JCastillo
' Propósito       :  Si hay una oferta establecer como actual. Devuelve un digito
' 0= Nada. 1= 2x1, 2=%, 3= a precio fijo
'---------------------------------------------------------------------------------------
'
Private Function Comprueba_Ofertas(conexion As ADODB.Connection) As Byte
Dim var As Variant

   'tipo de Oferta
   '0= 2x1, 1=%, 2= a precio fijo
   On Error GoTo Comprueba_Ofertas_Error
        
    var = devuelve_matriz("SELECT TIPO, DESCRIPCION, DCTO, IMPORTE FROM OFERTAS WHERE (FINICIO <= '" & Format(Date, "yyyymmdd") & "') AND (FFIN >= '" & Format(Date, "yyyymmdd") & "') AND (MBAJA = 0) AND (CODCAJA = " & CajaActual & ")", conexion)
    
    If Not IsArray(var) Then
        
        Comprueba_Ofertas = 0
        Exit Function
        
    Else
    
        OfertaDSC = "[" & Trim(var(1)) & "]"
        Comprueba_Ofertas = var(0)
        OfertaDcto = var(2)
        OfertaImp = var(3)
        
    End If

   On Error GoTo 0
   Exit Function

Comprueba_Ofertas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Comprueba_Ofertas de Formulario frmCabVen"
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : Inicializa_Ofertas
' Fecha/Hora     : 21/06/2004 11:08
' Autor             : JCastillo
' Propósito       : Subrutina donde se inician todos los procesos para gestionar ofertas.
'---------------------------------------------------------------------------------------
'
Public Sub Inicializa_Ofertas(conexion As ADODB.Connection)

   On Error GoTo Inicializa_Ofertas_Error

    'da las bajas de las ofertas que ya se han pasado de fecha ...
    Call Bajas_Ofertas(conexion)
    DoEvents
    OfertaActual = Comprueba_Ofertas(conexion)
    

   On Error GoTo 0
   Exit Sub

Inicializa_Ofertas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Inicializa_Ofertas de Módulo modOfertas"
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : Bajas_Ofertas
' Fecha/Hora    : 22/06/2004 16:53
' Autor         : JCastillo
' Propósito     :   da automaticamente de baja las ofertas
'---------------------------------- -----------------------------------------------------
'
Private Sub Bajas_Ofertas(conexion As ADODB.Connection)

   On Error GoTo Bajas_Ofertas_Error

    
    conexion.Execute "UPDATE OFERTAS SET MBAJA = -1 WHERE FFIN <= '" & Format(Date, "yyyymmdd") & "' AND CODCAJA = " & CajaActual
            

   On Error GoTo 0
   Exit Sub

Bajas_Ofertas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Bajas_Ofertas de Módulo modOfertas"
End Sub


