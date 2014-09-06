VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tekOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(ID As Long, TipoCpte As String, Emisor As Integer, Numero As Integer, RazonSocial As String, CUIT As String, Direccion As String, Fecha As String, CondIVA As String, PagoEfectivo As Double, Emitida As Boolean) As tekOrder
    'create a new object
    Dim objNewMember As tekOrder
    Set objNewMember = New tekOrder
    
    'set the properties passed into the method
    objNewMember.ID = ID
    objNewMember.TipoCpte = TipoCpte
    objNewMember.Numero = Numero
    objNewMember.Emisor = Emisor
    objNewMember.Fecha = Fecha
    objNewMember.RazonSocial = RazonSocial
    objNewMember.Direccion = Direccion
    objNewMember.CondIVA = CondIVA
    objNewMember.CUIT = CUIT
    objNewMember.PagoEfectivo = PagoEfectivo
    objNewMember.Emitida = Emitida


    mCol.Add objNewMember, CStr(objNewMember.ID)
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(Index As Variant) As tekOrder
  Set Item = mCol(Index)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(Index As Variant)
    mCol.Remove Index
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
