Attribute VB_Name = "modTekBiz"
Private Declare Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As _
    String) As Long

Private Declare Function WritePrivateProfileString Lib _
    "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Option Explicit


Public Const sVersion As String = "2.0.0"
Public bDevMode As Boolean
Public sPrinterPort As String

Public LogToFile As Boolean
Public Const sSystem As String = "Facturador"
Public Const FacturaB As String = "FCB"
Public Const FacturaA As String = "FCA"
Public Const NotaCA As String = "NCA"
Public Const NotaCB As String = "NCB"
Public Const Ticket As String = "TKB"



Public Sub ErrorHandler(Optional sAdditionalMessage As String)
    Dim sErr As String

    If Err.Number <> 0 Then
        sErr = Err.Source & ": " & vbCr & "(" & Err.Number & ") " & Err.Description & vbCr
    End If

'    If Not Connection Is Nothing Then
'        If rdoErrors.Count > 0 Then
'            Dim objHolder As Object
'            For Each objHolder In rdoErrors
'                sErr = sErr & objHolder.Source & ": " & vbCr & "(" & objHolder.Number & ") " & objHolder.Description & vbCr
'            Next
'        End If
'        rdoErrors.Clear
'    End If

    If sAdditionalMessage <> "" Then sErr = sErr & vbCr & vbCr & sAdditionalMessage

    ErrorLog sErr

    If Not sErr = "" Then
        Screen.MousePointer = vbDefault
        MsgBox sErr, vbCritical Or vbOKOnly
    End If

End Sub


Public Sub ErrorLog(sLine As String)
'    On Error Resume Next

    If Not LogToFile Then Exit Sub
    'If Len(sLine) > 512 Then sLine = Left(sLine, 508) & " ..."
    sLine = Replace(sLine, vbCr, " ")
    sLine = Replace(sLine, vbLf, "")
    Open App.Path & "\Error.log" For Append As #1
    Print #1, Format(Now, "YYYYMMDD hh:mm:ss") & "-" & App.EXEName & "-" & App.Major & "." & App.Minor & "." & App.Revision & "-" & App.ThreadID & "| " & sLine
    Close #1

End Sub


Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    ' help from MSXML
    Set objXML = New MSXML2.DOMDocument

    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

     ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing

End Function

Function DecodeBase64(ByVal strData As String) As Byte()


    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    ' help from MSXML
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue

    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing

End Function


Function sBase64Enc(sText As String)
    sBase64Enc = EncodeBase64(StrConv(sText, vbFromUnicode))
End Function

    
' return the text contained in an INI file key
Public Function ReadINI(strSection As String, strKeyName As String, _
        ByVal strFileName As String, Optional lStringBuffer As Long = 255, _
        Optional iniDirPathName As String) As String

' WARNING: this functions uses a limited buffer, specified by
' the instruction strText = Space(buffer).
' Every exceeding char will be TRIMMED out. If you think this
' could cut off a part of the retrieved data, you can increase
' the argument in the Space() function.
    
    Dim intLen As Long
    Dim strText As String, strIniFile As String
    ' set INI file default directory
    If iniDirPathName = "" Then iniDirPathName = App.Path
    ' cut off the final backslash char from path
    If Right$(iniDirPathName, 1) = "\" Then
        Mid$(iniDirPathName, Len(iniDirPathName), 1) = Space$(1)
        iniDirPathName = RTrim$(iniDirPathName)
    End If
    ' build INI file's complete path
    strIniFile = iniDirPathName & "\" & strFileName
    ' read data in the file, and check data for errors
    strText = Space(lStringBuffer) ' BUFFER
    intLen = GetPrivateProfileString(strSection, strKeyName, "", _
        strText, Len(strText), strIniFile)
    If intLen > -1 Then
        strText = Left(strText, intLen)
    Else
        MsgBox "Error into INI file"
        Exit Function
    End If
    ReadINI = strText
' USAGE: MyProperty = ReadINI(Section, KeyName, INIFileName, buffer)
' SAMPLE: Me.FontName = ReadINI("Font", "FontName", "Banner.ini", 255)
End Function

