VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tekOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lID As Long
Private sTipoCpte As String
Private iEmisor As Integer
Private iNumero As Long
Private sFecha As String
Private sRazonSocial As String
Private sDireccion As String
Private sCondIVA As String
Private sCUIT As String
Private sCpteRel As String
Private dPagoEfectivo As Double
Private bEmitida As Boolean

' esta clase a su vez tiene una coleccion de items
Private mCol As Collection
Private oItems As tekOrderItems


Public Property Get ID() As Long
   ID = lID
End Property

Public Property Let ID(ByVal NewValue As Long)
   lID = NewValue
End Property

Public Property Get TipoCpte() As String
   TipoCpte = sTipoCpte
End Property

Public Property Let TipoCpte(ByVal NewValue As String)
   sTipoCpte = NewValue
End Property

Public Property Get Emisor() As Integer
   Emisor = iEmisor
End Property

Public Property Let Emisor(ByVal NewValue As Integer)
   iEmisor = NewValue
End Property

Public Property Get Numero() As Long
   Numero = iNumero
End Property

Public Property Let Numero(ByVal NewValue As Long)
   iNumero = NewValue
End Property

Public Property Get Fecha() As String
   Fecha = sFecha
End Property

Public Property Let Fecha(ByVal NewValue As String)
   sFecha = NewValue
End Property

Public Property Get RazonSocial() As String
   RazonSocial = sRazonSocial
End Property

Public Property Let RazonSocial(ByVal NewValue As String)
   sRazonSocial = NewValue
End Property

Public Property Get Direccion() As String
   Direccion = sDireccion
End Property

Public Property Let Direccion(ByVal NewValue As String)
   sDireccion = NewValue
End Property

Public Property Get CondIVA() As String
   CondIVA = sCondIVA
End Property

Public Property Let CondIVA(ByVal NewValue As String)
   sCondIVA = NewValue
End Property

Public Property Get CUIT() As String
   CUIT = sCUIT
End Property

Public Property Let CUIT(ByVal NewValue As String)
   sCUIT = NewValue
End Property

Public Property Get CpteRel() As String
   CpteRel = sCpteRel
End Property

Public Property Let CpteRel(ByVal NewValue As String)
   sCpteRel = NewValue
End Property


Public Property Get PagoEfectivo() As Double
   PagoEfectivo = dPagoEfectivo
End Property

Public Property Let PagoEfectivo(ByVal NewValue As Double)
   dPagoEfectivo = NewValue
End Property

Public Property Get Emitida() As Boolean
   Emitida = bEmitida
End Property

Public Property Let Emitida(ByVal NewValue As Boolean)
   bEmitida = NewValue
End Property



' ---------------------------------------------- FUNCIONES PARA LA COLECCION DE ITEMS
Public Property Get OrderItems() As tekOrderItems
   Set OrderItems = oItems
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    Set oItems = New tekOrderItems
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    Set oItems = Nothing
End Sub
