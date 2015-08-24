VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tekOrderItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(ItemId As Integer, Descripcion As String, Cantidad As Integer, Precio As Double, IVA As Double, ImpInt As Double, Despacho As String, Desc As Double) As tekOrderItem
    'create a new object
    Dim objNewMember As tekOrderItem
    Set objNewMember = New tekOrderItem

    'set the properties passed into the method
    objNewMember.ItemId = ItemId
    objNewMember.Descripcion = Descripcion
    objNewMember.Cantidad = Cantidad
    objNewMember.Precio = Precio
    objNewMember.IVA = IVA
    objNewMember.ImpInt = ImpInt
    objNewMember.NroDespacho = Despacho
    objNewMember.Desc = Desc
    
    mCol.Add objNewMember, CStr(objNewMember.ItemId)
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(Index As Variant) As tekOrderItem
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


