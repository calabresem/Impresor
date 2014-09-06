VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------
' Clase de comunicacion con el WS - TEKAR 2005
' --------------------

' miscelaneos
Private sVersion As String
Private sLanguage As String
Private sRootElement As String

' propiedades
Private sURL As String
Private sLogin As String
Private sPassword As String
Private bError As Boolean
Private sErrorNumber As String
Private sErrorDescription As String
Private sRequest As String
Private sResponse As String
Private sAPIError As String
Private sBasicAuthLogin As String
Private sBasicAuthPassword As String
Private bAlreadyHashed As Boolean

' objetos
Private oMD5 As MD5


Public Property Get URL() As String
    URL = sURL
End Property
Public Property Let URL(ByVal vNewValue As String)
    sURL = vNewValue
End Property

Public Property Get RootElement() As String
    RootElement = sRootElement
End Property
Public Property Let RootElement(ByVal vNewValue As String)
    sRootElement = vNewValue
End Property

Public Property Get AlreadyHashed() As Boolean
    AlreadyHashed = bAlreadyHashed
End Property
Public Property Let AlreadyHashed(ByVal vNewValue As Boolean)
    bAlreadyHashed = vNewValue
End Property

Public Property Let AuthLogin(ByVal vNewValue As String)
    sBasicAuthLogin = vNewValue
End Property

Public Property Let AuthPassword(ByVal vNewValue As String)
    sBasicAuthPassword = vNewValue
End Property

Public Property Let Login(ByVal vNewValue As String)
    sLogin = vNewValue
End Property

Public Property Let Password(ByVal vNewValue As String)
    Set oMD5 = New MD5
    ' si el string ya viene en MD5 hay que avisar
    If bAlreadyHashed Then
        sPassword = vNewValue
    Else
        sPassword = oMD5.DigestStrToHexStr(vNewValue)
    End If
    Set oMD5 = Nothing
End Property

Public Property Get Error() As Boolean
    Error = bError
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = sErrorDescription
End Property

Public Property Get APIError() As String
    APIError = sAPIError
End Property

Public Property Get ErrorNumber() As String
    ErrorNumber = sErrorNumber
End Property

Public Property Get Request() As String
    Request = sRequest
End Property

Public Property Get Response() As String
    Response = sResponse
End Property

' --------------------------------------------------------------------
' Inicializa la clase
' Autor: Marcos Calabrese (TEKAR - 11/2005)
' ---
Private Sub Class_Initialize()

    sVersion = "1.0"
    sLanguage = "en"
    sRootElement = "puzzlemania"

End Sub


' --------------------------------------------------------------------
' Arma el XML y se conecta al webservice
' Autor: Marcos Calabrese (TEKAR - 11/2005)
' ---
Public Function ExecuteAPI(sData As String) As Boolean

Dim objHTTP As New MSXML2.XMLHTTP40
Dim strReturn As String
Dim objReturn As MSXML2.DOMDocument40
Dim dblTax As Double
Dim strQuery As String

'On Error GoTo ErrHandler

    Set objReturn = New MSXML2.DOMDocument40

    sRequest = sData

    ' Set up to post to our local server
    objHTTP.Open "POST", sURL, False
    
    ' Set a standard SOAP/ XML header for the content-type
    objHTTP.setRequestHeader "Content-Type", "text/xml"
    
    ' agrega al header tipo de autenticacion basica
    objHTTP.setRequestHeader "Authorization", "Basic " & sBase64Enc(sBasicAuthLogin & ":" & sBasicAuthPassword)
    
    
    ' envia el XML
    objHTTP.send sData
    
    ' obtiene la respuesta
    strReturn = objHTTP.responseText


    ' hacemos publica la respuesta
    sResponse = strReturn


    ' carga el XML en formato DOM
    objReturn.loadXML strReturn

    ' hubo error de parseo?
    If objReturn.parseError <> 0 Then

        If objReturn.parseError = -1072897501 Then
            sAPIError = "Validate failed because the document does not contain exactly one root node."
        ElseIf objReturn.parseError = -1072896760 Then
            sAPIError = "An invalid character was found in text content."
        ElseIf objReturn.parseError = -1072896683 Then
            sAPIError = "Only one top level element is allowed in an XML document."
        ElseIf objReturn.parseError = -1072896659 Then
            sAPIError = "There is an unclosed element in the XML document."
        ElseIf objReturn.parseError = -1072896759 Then
            sAPIError = "Required white space was missing."
        Else
            sAPIError = "Error: " & objReturn.parseError
        End If
        
       
        ExecuteAPI = False
        Exit Function
        
    End If


    ' buscamos el item de error
    If objReturn.selectSingleNode(sRootElement & "/result").Text = "ERROR" Then
        
        ' si hubo error, marcamos el flag y mostramos la descripcion
        bError = True
        sErrorNumber = objReturn.selectSingleNode(sRootElement & "/error/number").Text
        sErrorDescription = objReturn.selectSingleNode(sRootElement & "/error/description").Text
   
    End If
   
    
    Set objReturn = Nothing
    
    
    ExecuteAPI = True
 
    Exit Function

ErrHandler:
    ErrorHandler

End Function


' --------------------------------------------------------------------
' Genera el elemento principal - el root element
' Autor: Marcos Calabrese (TEKAR - 11/2005)
' ---
Private Function XMLHeader(sXML As String) As String

Dim sBuffer As String

'On Error GoTo ErrHandler

    sBuffer = "<" & sRootElement & " version=""" & sVersion & """ lang=""" & sLanguage & """ output=""text""  >"
    
    ' agregamos el elemento de auth
    sBuffer = sBuffer & XMLAuthElement
    
    ' ahora el XML que viene x parametro
    sBuffer = sBuffer & sXML
    
    sBuffer = sBuffer & "</" & sRootElement & ">"

    XMLHeader = sBuffer
    
    Exit Function

ErrHandler:
    ErrorHandler

End Function

' --------------------------------------------------------------------
' Genera el elemento de acceso
' Autor: Marcos Calabrese (TEKAR - 11/2005)
' ---
Private Function XMLAuthElement() As String

Dim sBuffer As String

'On Error GoTo ErrHandler

    sBuffer = "<auth login=""" & sLogin & """ password=""" & sPassword & """ />"

    XMLAuthElement = sBuffer

    Exit Function

ErrHandler:
    ErrorHandler

End Function


' --------------------------------------------------------------------
' Funcion para probar el webservice
' Autor: Marcos Calabrese (TEKAR - 11/2005)
' ---
Public Function Ping() As Boolean

Dim sBuffer As String

'On Error GoTo ErrHandler

    ' prepara el comando
    sBuffer = sBuffer & "<command>Ping</command>" & sData
    
    ' le da formato al XML
    sBuffer = XMLHeader(sBuffer)
    
    ' ejecutamos el API
    Ping = ExecuteAPI(sBuffer)
    


    Exit Function

ErrHandler:
    ErrorHandler
    Ping = False

End Function


' --------------------------------------------------------------------
' Funcion para consultar facturas pendientes de emision
' Autor: Marcos Calabrese (TEKAR - 11/2005)
' ---
Public Function GetPendingOrders(Optional sData As String) As Boolean

Dim sBuffer As String

'On Error GoTo ErrHandler

    ' prepara el comando
    sBuffer = sBuffer & "<command>GetPendingInvoices</command>"
    sBuffer = sBuffer & sData
    
    ' le da formato al XML
    sBuffer = XMLHeader(sBuffer)
    
    ' ejecutamos el API
    GetPendingOrders = ExecuteAPI(sBuffer)


    Exit Function

ErrHandler:
    ErrorHandler
    GetPendingOrders = False

End Function


' --------------------------------------------------------------------
' Funcion para marcar facturas impresas
' Autor: Marcos Calabrese (TEKAR - 11/2005)
' ---
Public Function SetPrintedInvoices(Optional sData As String) As Boolean

Dim sBuffer As String

'On Error GoTo ErrHandler

    ' prepara el comando
    sBuffer = sBuffer & "<command>SetPrintedInvoices</command>"
    sBuffer = sBuffer & sData
    
    ' le da formato al XML
    sBuffer = XMLHeader(sBuffer)
    
    ' ejecutamos el API
    SetPrintedInvoices = ExecuteAPI(sBuffer)


    Exit Function

ErrHandler:
    ErrorHandler
    SetPrintedInvoices = False

End Function

