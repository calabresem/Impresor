VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tekOrderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iItemID As Integer
Private sDescripcion As String
Private iCantidad As Integer
Private dPrecio As Double
Private dIVA As Double
Private dImpInt As Double
Private sNroDespacho As String


' ----------------- ID del item
Public Property Get ItemId() As Integer
   ItemId = iItemID
End Property

Public Property Let ItemId(ByVal NewValue As Integer)
   iItemID = NewValue
End Property

' ----------------- descripcion
Public Property Get Descripcion() As String
   Descripcion = sDescripcion
End Property

Public Property Let Descripcion(ByVal NewValue As String)
   sDescripcion = NewValue
End Property

' ----------------- cantidad
Public Property Get Cantidad() As Integer
   Cantidad = iCantidad
End Property

Public Property Let Cantidad(ByVal NewValue As Integer)
   iCantidad = NewValue
End Property

' ----------------- precio unitario
Public Property Get Precio() As Double
   Precio = dPrecio
End Property

Public Property Let Precio(ByVal NewValue As Double)
   dPrecio = NewValue
End Property

' ----------------- % de IVA
Public Property Get IVA() As Double
   IVA = dIVA
End Property

Public Property Let IVA(ByVal NewValue As Double)
   dIVA = NewValue
End Property

' ----------------- % de Impuesto Interno
Public Property Get ImpInt() As Double
   ImpInt = dImpInt
End Property

Public Property Let ImpInt(ByVal NewValue As Double)
   dImpInt = NewValue
End Property

' ----------------- nro despacho
Public Property Get NroDespacho() As String
   NroDespacho = sNroDespacho
End Property

Public Property Let NroDespacho(ByVal NewValue As String)
   sNroDespacho = NewValue
End Property

