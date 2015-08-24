VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTekBiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const sDSN As String = "PUZZLE_TEKBIZ"
Private iOrdersProcessed As Integer
Private sLastStatus As String
Private Const iEmisor As Integer = 2
Private oOrders As tekOrders
'Private oAPI As clsAPI
Private bError As Boolean
Private sErrDescription As String
Private Const sRootElement As String = "tekarbiz"

' ------------------------------------------------
' PROPIEDADES
' ------------------------------------------------
Public Property Get Status() As String
    Status = sLastStatus
End Property

Public Property Get OrdersProcessed() As Integer
    OrdersProcessed = iOrdersProcessed
End Property

Public Property Get Orders() As tekOrders
   Set Orders = oOrders
End Property

Public Property Get API() As clsAPI
    Set API = oAPI
End Property

Public Property Get Error() As Boolean
    Error = bError
End Property

Public Property Get ErrDescription() As String
    ErrDescription = sErrDescription
End Property


' ------------------------------------------------
' FUNCIONES
' ------------------------------------------------


' ------------------------------------------------
Public Function Initialize() As Boolean

    'On Error GoTo ErrHandler


    ' --------------------------------------------------------------- config.xml
    ' instanciamos los objetos
    If IsObject(oOrders) = True And Not IsEmpty(oOrders) Then
        Set oOrders = Nothing
    End If
    
    ' instanciamos los objetos
    Set oOrders = New tekOrders


    Initialize = True

    Exit Function

ErrHandler:
    ErrorHandler
End Function

' ------------------------------------------------
Public Function Terminate() As Boolean

    'On Error GoTo ErrHandler

    Set oAPI = Nothing

    Terminate = True

    Exit Function

ErrHandler:
    ErrorHandler
End Function


' ------------------------------------------------
Public Function Restart() As Boolean

    'On Error GoTo ErrHandler

    ' reinicia las ordenes
    Set oOrders = Nothing
    Set oOrders = New tekOrders


    Exit Function

ErrHandler:
    ErrorHandler
End Function


' ------------------------------------------------
' USANDO WEBSERVICES
Public Function GetPendingOrders(ByVal sXML As String) As Integer

Dim oResponse As MSXML2.DOMDocument30
Dim oOrdersNode As MSXML2.IXMLDOMNodeList
Dim oItemsNode As MSXML2.IXMLDOMNodeList
'Dim sXML As String
Dim x%, i%, j%
Dim oOrder As tekOrder, oItem As tekOrderItem


    bError = False
    sErrDescription = ""


    'On Error GoTo ErrHandler
    x = 0
    
    sLastStatus = "OK"
    
    'Debug.Print oAPI.Response
    
    ' levantamos la respuesta como XML
    Set oResponse = New MSXML2.DOMDocument30
    oResponse.loadXML sXML

    
    ' obtenemos la cantidad de ordenes para procesar
    x = oResponse.selectNodes(sRootElement & "/pending_invoices/invoices/invoice").Length
    
   
    'Exit Function
    Dim lID
    
    ' revisamos las pendientes (no emitidas) para el emisor en uso
    If x > 0 Then
    
       
        ' referencia los items
        Set oOrdersNode = oResponse.selectNodes(sRootElement & "/pending_invoices/invoices/invoice")

        ' el XML es 0-based
        For i = 0 To (x - 1)

            ' agregamos la orden
            Set oOrder = oOrders.Add(CLng(oOrdersNode.Item(i).selectSingleNode("@id").Text), oOrdersNode.Item(i).selectSingleNode("@type").Text, CInt(oOrdersNode.Item(i).selectSingleNode("@source").Text), CInt(oOrdersNode.Item(i).selectSingleNode("@number").Text), oOrdersNode.Item(i).selectSingleNode("customer/name").Text, oOrdersNode.Item(i).selectSingleNode("customer/@tax_id").Text, oOrdersNode.Item(i).selectSingleNode("customer/address").Text, oOrdersNode.Item(i).selectSingleNode("@date").Text, oOrdersNode.Item(i).selectSingleNode("customer/@tax_type").Text, CDbl(oOrdersNode.Item(i).selectSingleNode("payment").Text), False)

            ' seleccionamos los items de la orden
            Set oItemsNode = oOrdersNode.Item(i).selectNodes("items/item")
            
            For j = 0 To oItemsNode.Length - 1
            
                ' y los agregamos
                Set oItem = oOrders.Item(oOrders.Count).OrderItems.Add(j + 1, oItemsNode.Item(j).selectSingleNode("description").Text, CInt(oItemsNode.Item(j).selectSingleNode("@quantity").Text), CDbl(oItemsNode.Item(j).selectSingleNode("@price").Text), CDbl(oItemsNode.Item(j).selectSingleNode("@tax").Text), 0, "", CDbl(oItemsNode.Item(j).selectSingleNode("@discount").Text))
            
            Next j

        Next i
    
        GetPendingOrders = x
        sLastStatus = "GET_PENDINGS_OK"

    Else
        
        GetPendingOrders = 0
        sLastStatus = "GET_PENDINGS_OK"

    End If
    
    
    
    iOrdersProcessed = x
    GetPendingOrders = x

    Exit Function

ErrHandler:
    ErrorHandler
End Function

' ------------------------------------------------
' HACE UN PING AL WS
Function CheckWebService() As Boolean

    'DoEvents
    CheckWebService = oAPI.Ping
    
    'Debug.Print oAPI.Request
    'Debug.Print oAPI.Response
    
    If CheckWebService = False Then
        'Debug.Print "Error: " & oAPI.APIError
        sErrDescription = oAPI.APIError
        bError = True
    End If

End Function


' ------------------------------------------------
' esta funcion recorre las ordenes cargadas y las marca como impresas
Public Function SetPrintedOrders() As Boolean

    Dim sXML As String, sIDs As String
    Dim i%
    
    'On Error GoTo ErrHandler

    
    If oOrders.Count > 0 Then

        sXML = "<set_printed_invoices>"

        For i = 1 To oOrders.Count
        
            sXML = sXML & "<invoice id=""" & oOrders.Item(i).ID & """ />"
       
        Next i
    
        sXML = sXML & "</set_printed_invoices>"
    
    End If
    
    
    ' llamamos al web service
    If Not oAPI.SetPrintedInvoices(sXML) Then
        sLastStatus = "SET_PRINTED_ERROR"
        sErrDescription = oAPI.APIError
        bError = True
        SetPrintedOrders = x
        Exit Function
    End If

    ' la respuesta nos muestra un error?
    If oAPI.Error Then
        sLastStatus = "SET_PRINTED_ERROR"
        sErrDescription = "[" & oAPI.ErrorNumber & "] " & oAPI.ErrorDescription
        bError = True
        SetPrintedOrders = x
        Exit Function
    End If
    
    
   
    SetPrintedOrders = True

    Exit Function

ErrHandler:
    ErrorHandler

End Function

Private Sub Class_Terminate()
    Set oOrders = Nothing
'    Set oAPI = Nothing
End Sub

