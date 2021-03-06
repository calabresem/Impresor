VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{39ABE45D-F077-4D34-A361-6906C77D67F7}#1.0#0"; "Fiscal150423.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturador"
   ClientHeight    =   2055
   ClientLeft      =   2985
   ClientTop       =   3360
   ClientWidth     =   4605
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4605
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   2760
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame frmStatus 
      Caption         =   "Estado"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtStatus 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Timer tmrOrders 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1920
      Top             =   2760
   End
   Begin FiscalPrinterLibCtl.HASAR HASAR1 
      Left            =   960
      OleObjectBlob   =   "frmMain.frx":1242
      Top             =   2880
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Menu"
      Begin VB.Menu mnuShow 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReporteX 
         Caption         =   "Reporte X"
      End
      Begin VB.Menu mnuReporteZ 
         Caption         =   "Reporte Z"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuError 
      Caption         =   "Sistema"
      Begin VB.Menu mnuSysDebug 
         Caption         =   "Debug"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysCancelAll 
         Caption         =   "Cancelar cptes. abiertos"
      End
      Begin VB.Menu mnuSysReprint 
         Caption         =   "Reimprimir el ultimo comprobante"
      End
   End
   Begin VB.Menu mnuMinimize 
      Caption         =   "Minimizar"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------
' COPYRIGHT 2005, TEKAR
' Autor: Marcos Calabrese
' ------------------------------------------------

Private oTekBiz As clsTekBiz
Private bPrinting As Boolean
Public FS As String

'Separates actual packets (stands for 'End of Packet').
Private Const EOP As String = "???"
Private strBuffer As String 'Data buffer.

Private Function NumeroInventado() As Long

    Randomize
    NumeroInventado = Int((99999 * Rnd()) + 999)

End Function



Private Sub Form_Load()

    ' Inicializamos todo
    Init
    
    ' minimizamos auto
    'Call mnuMinimize_Click
    
    MinimizeForm


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim respbox As Integer

    respbox = MsgBox("Esta seguro que desea cerrar la aplicacion?", vbQuestion + vbYesNo, "  SALIDA !!!")
    
    If (respbox = vbNo) Then
        Cancel = True
    Else
        TerminateApp
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TerminateApp

End Sub

Private Function ImprimirTickets(ByRef oTekBiz As clsTekBiz)

Dim x As Integer, Y As Integer
Dim sComando As String, sRazonSocial As String, sDomicilio As String, sCUIT As String
Dim TipoDoc As TiposDeDocumento
Dim CondImp As TiposDeResponsabilidades
Dim sEmisorFMT As String, sNumeroFMT As String
Dim cTotDesc As Double
      
    ' hay algo para imprimir? (se supone que si, pero...)
    If oTekBiz.Orders.Count = 0 Then
        Exit Function
    End If


'On Error GoTo impresora_apag
       
Procesar:
    
    Screen.MousePointer = vbHourglass
    
    For x = 1 To oTekBiz.Orders.Count
    
        ' -----------------
        sEmisorFMT = Format(oTekBiz.Orders.Item(x).Emisor, "0000")
        sNumeroFMT = Format(oTekBiz.Orders.Item(x).Numero, "00000000")
        cTotDesc = 0
        
        ' ----------------- seleccionar tipo de comprobante
        Select Case oTekBiz.Orders.Item(x).TipoCpte

            ' ------------------------------ TICKET FACTURA A
            Case FacturaA

                '// ----------------------------------------------
                'sRazonSocial = Left(oTekBiz.Orders.Item(x).RazonSocial, 30)
                sRazonSocial = oTekBiz.Orders.Item(x).RazonSocial
                'sDomicilio = Left(oTekBiz.Orders.Item(x).Direccion, 40)
                sDomicilio = oTekBiz.Orders.Item(x).Direccion
                ' le sacamos los guiones al CUIT
                sCUIT = Replace(oTekBiz.Orders.Item(x).CUIT, "-", "")
                
                ' completamos con espacios
                'sRazonSocial = sRazonSocial & String((30 - Len(sRazonSocial)), " ")
                'sDomicilio = sDomicilio & String((40 - Len(sDomicilio)), " ")
                'sDomicilio = String((40 - Len("")), " ")
                
                If bDevMode = False Then
                    HASAR1.DatosCliente sRazonSocial, sCUIT, TIPO_CUIT, RESPONSABLE_INSCRIPTO, sDomicilio
                    HASAR1.AbrirComprobanteFiscal TICKET_FACTURA_A
                    
                Else
                    Debug.Print "Razon social: " & sRazonSocial & " - CUIT: " & sCUIT & " - Domicilio: " & sDomicilio
                    
                End If
                
                
                '------------------ IMPRIME ITEMS
                For Y = 1 To oTekBiz.Orders.Item(x).OrderItems.Count
           
                    ' imprimimos el detalle del ticket
                    If bDevMode = False Then
                        HASAR1.ImprimirItem oTekBiz.Orders.Item(x).OrderItems.Item(Y).Descripcion, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Cantidad, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Precio, oTekBiz.Orders.Item(x).OrderItems.Item(Y).IVA, oTekBiz.Orders.Item(x).OrderItems.Item(Y).ImpInt
                        
                        If oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc <> 0 Then
                            cTotDesc = cTotDesc + oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc
                        End If

                    Else
                        Debug.Print "Imprime el detalle: " & oTekBiz.Orders.Item(x).OrderItems.Item(Y).Descripcion & _
                        " - " & CStr(oTekBiz.Orders.Item(x).OrderItems.Item(Y).Cantidad) & " x " & CStr(oTekBiz.Orders.Item(x).OrderItems.Item(Y).Precio)
                    End If
     
                Next Y


                ' Imprime el descuento
                If cTotDesc <> 0 Then
                    HASAR1.DescuentoGeneral "Descuento", cTotDesc, True
                End If
                
                
                '------------------ IMPRIME PAGO
                If bDevMode = False Then
                    If oTekBiz.Orders.Item(x).PagoEfectivo <> 0 Then
                        HASAR1.ImprimirPago "Efectivo", oTekBiz.Orders.Item(x).PagoEfectivo
                    End If

                    ' cierra el comprobante
                    HASAR1.CerrarComprobanteFiscal
                    
                    ' Le da el numero que le dio la impresora
                    oTekBiz.Orders.Item(x).Numero = HASAR1.UltimoDocumentoFiscalA

                Else
                    Debug.Print "Pago:  " & CStr(oTekBiz.Orders.Item(x).PagoEfectivo)
                    oTekBiz.Orders.Item(x).Numero = NumeroInventado()
                    
                End If
                


            ' ------------------------------ TICKET FACTURA B/C
            Case Ticket
            Case FacturaB

                ' Es consumidor final, salvo que sea EXENTO
                If oTekBiz.Orders.Item(x).CondIVA = "E" Then
                    sRazonSocial = oTekBiz.Orders.Item(x).RazonSocial
                    CondImp = RESPONSABLE_EXENTO
                    sCUIT = Replace(oTekBiz.Orders.Item(x).CUIT, "-", "")
                    TipoDoc = TIPO_CUIT
                    
                Else
                    sRazonSocial = "CONSUMIDOR FINAL"
                    CondImp = CONSUMIDOR_FINAL
                    sCUIT = "00000000"
                    TipoDoc = TIPO_DNI
                    
                End If
                

                If bDevMode = False Then
                    HASAR1.DatosCliente sRazonSocial, sCUIT, TipoDoc, CondImp, "."
                    HASAR1.AbrirComprobanteFiscal TICKET_FACTURA_B
                Else
                    Debug.Print "Abre comprobante: TICKET FACTURA B"
                End If


                '------------------ IMPRIME ITEMS
                For Y = 1 To oTekBiz.Orders.Item(x).OrderItems.Count

                    If bDevMode = False Then
                        HASAR1.ImprimirItem oTekBiz.Orders.Item(x).OrderItems.Item(Y).Descripcion, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Cantidad, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Precio, oTekBiz.Orders.Item(x).OrderItems.Item(Y).IVA, oTekBiz.Orders.Item(x).OrderItems.Item(Y).ImpInt

                        If oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc <> 0 Then
                            cTotDesc = cTotDesc + oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc
                        End If
                        
                    Else
                        Debug.Print "Imprime el detalle"
                    End If

                Next Y

                
                ' Imprime el descuento
                If cTotDesc <> 0 Then
                    HASAR1.DescuentoGeneral "Descuento", cTotDesc, True
                End If

                
                '------------------ IMPRIME ITEMS
                If bDevMode = False Then
                    If oTekBiz.Orders.Item(x).PagoEfectivo <> 0 Then
                        HASAR1.ImprimirPago "Efectivo", oTekBiz.Orders.Item(x).PagoEfectivo
                    End If

                    ' cierra comprobante
                    HASAR1.CerrarComprobanteFiscal
                    
                    ' Le da el numero que le dio la impresora
                    oTekBiz.Orders.Item(x).Numero = HASAR1.UltimoDocumentoFiscalBC

                Else
                    Debug.Print "Imprime el pago y cierre"
                    oTekBiz.Orders.Item(x).Numero = NumeroInventado()
                    
                End If



            ' ------------------------------ NOTA CREDITO A
            Case NotaCA

                If bDevMode = False Then
    
                    '// ----------------------------------------------
                    ' Prepara los datos de cabecera
                    sRazonSocial = Left(oTekBiz.Orders.Item(x).RazonSocial, 30)
                    sDomicilio = Left(oTekBiz.Orders.Item(x).Direccion, 40)
                    sCUIT = Replace(oTekBiz.Orders.Item(x).CUIT, "-", "")
                    sRazonSocial = sRazonSocial & String((30 - Len(sRazonSocial)), " ")
                    sDomicilio = sDomicilio & String((40 - Len(sDomicilio)), " ")
                    
                    ' Informa el comprobante relacionado
                    HASAR1.Enviar Chr$(147) & FS & "1" & FS & oTekBiz.Orders.Item(x).CpteRel
                    
                    ' Informa la cabecera
                    comando = Chr$(98) & FS & sRazonSocial & FS & sCUIT & FS & "I" & FS & "C" & _
                              FS & "."
                    HASAR1.Enviar comando
                    HASAR1.Enviar Chr$(128) & FS & "R" & FS & "T"
    
                Else
                    Debug.Print "Imprime encabezado."
                End If

                
                For Y = 1 To oTekBiz.Orders.Item(x).OrderItems.Count
            
                    If bDevMode = False Then
                        HASAR1.ImprimirItem oTekBiz.Orders.Item(x).OrderItems.Item(Y).Descripcion, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Cantidad, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Precio, oTekBiz.Orders.Item(x).OrderItems.Item(Y).IVA, oTekBiz.Orders.Item(x).OrderItems.Item(Y).ImpInt
                        
                        If oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc <> 0 Then
                            cTotDesc = cTotDesc + oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc
                        End If
                        
                    Else
                        Debug.Print "Imprime detalle: " & oTekBiz.Orders.Item(x).OrderItems.Item(Y).Descripcion
                    End If
                                
                Next Y
                
                
                ' Imprime el descuento
                If cTotDesc <> 0 Then
                    HASAR1.DescuentoGeneral "Descuento", cTotDesc, True
                End If
                

                If bDevMode = False Then
                    If oTekBiz.Orders.Item(x).PagoEfectivo <> 0 Then
                        HASAR1.ImprimirPago "Efectivo", oTekBiz.Orders.Item(x).PagoEfectivo
                    End If

                    '//---------------------------------------------------------------
                    HASAR1.CerrarDNFH
                    
                    ' Le da el numero que le dio la impresora
                    oTekBiz.Orders.Item(x).Numero = HASAR1.UltimaNotaCreditoA
                    
                Else
                    Debug.Print "Imprime pago y cierra."
                    oTekBiz.Orders.Item(x).Numero = NumeroInventado()
                    
                End If



            ' ------------------------------ NOTA CREDITO B/C
            Case NotaCB

                If bDevMode = False Then
                
                    Dim sCondImp As String, sTipoDoc As String


                    ' Es consumidor final, salvo que sea EXENTO
                    If oTekBiz.Orders.Item(x).CondIVA = "E" Then
                        sRazonSocial = oTekBiz.Orders.Item(x).RazonSocial
                        CondImp = RESPONSABLE_EXENTO
                        sCondImp = "E"
                        sCUIT = Replace(oTekBiz.Orders.Item(x).CUIT, "-", "")
                        TipoDoc = TIPO_CUIT
                        sTipoDoc = "C"
                        
                    Else
                        sRazonSocial = "CONSUMIDOR FINAL"
                        CondImp = CONSUMIDOR_FINAL
                        sCondImp = "C"
                        sCUIT = "9999"
                        sTipoDoc = "2"
                        
                    End If

                    
                    ' Informa la cabecera
                    HASAR1.Enviar Chr$(147) & FS & "1" & FS & oTekBiz.Orders.Item(x).CpteRel
                    comando = Chr$(98) & FS & sRazonSocial & FS & sCUIT & FS & sCondImp & FS & sTipoDoc & _
                              FS & "."
                    HASAR1.Enviar comando
                    HASAR1.Enviar Chr$(128) & FS & "S" & FS & "T"
                    
                Else
                    Debug.Print "Imprime encabezado."
                End If


                For Y = 1 To oTekBiz.Orders.Item(x).OrderItems.Count
            
                    If bDevMode = False Then
                        HASAR1.ImprimirItem oTekBiz.Orders.Item(x).OrderItems.Item(Y).Descripcion, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Cantidad, oTekBiz.Orders.Item(x).OrderItems.Item(Y).Precio, oTekBiz.Orders.Item(x).OrderItems.Item(Y).IVA, oTekBiz.Orders.Item(x).OrderItems.Item(Y).ImpInt

                        If oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc <> 0 Then
                            cTotDesc = cTotDesc + oTekBiz.Orders.Item(x).OrderItems.Item(Y).Desc
                        End If

                    Else
                        Debug.Print "Imprime detalle."
                    End If
                    
                Next Y
                
                ' Imprime el descuento
                If cTotDesc <> 0 Then
                    HASAR1.DescuentoGeneral "Descuento", cTotDesc, True
                End If
                

                If bDevMode = False Then
                    If oTekBiz.Orders.Item(x).PagoEfectivo <> 0 Then
                        HASAR1.ImprimirPago "Efectivo", oTekBiz.Orders.Item(x).PagoEfectivo
                    End If
                
                    '//---------------------------------------------------------------
                    HASAR1.CerrarDNFH

                    ' Le da el numero que le dio la impresora
                    oTekBiz.Orders.Item(x).Numero = HASAR1.UltimaNotaCreditoBC

                Else
                    Debug.Print "Imprime detalle."
                    oTekBiz.Orders.Item(x).Numero = NumeroInventado()
                End If
                
                
        Case Else
            Debug.Print "El comprobante no corresponde."
            

       
       End Select
   
   
    Next x
    
    ' puntero
    Screen.MousePointer = vbNormal

    Exit Function

impresora_apag:

    Screen.MousePointer = vbNormal
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Function

Private Sub mnuSysCancelAll_Click()

    On Error GoTo impresora_apag
    
    
Procesar:
    
    Screen.MousePointer = vbHourglass
    
    
    If bDevMode = False Then
        ' envia el comando de cancelacion
        HASAR1.Enviar Chr$(152)
    Else
        Debug.Print "Cancela todos los comprobantes."
    End If

    Screen.MousePointer = vbNormal

    Exit Sub

impresora_apag:

    Screen.MousePointer = vbNormal
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Private Sub mnuExit_Click()

    ' terminamos
    TerminateApp

End Sub

Private Sub mnuMinimize_Click()
    ' agrega el icono en el systray
    TrayAdd hwnd, Me.Icon, sSystem, MouseMove
    mnuHide_Click
End Sub

' --------------------------------------------------------------------------
' REPORTES
' --------------------------------------------------------------------------
Private Sub mnuReporteX_Click()
    
'    On Error GoTo impresora_apag
    
Procesar:
    
    Screen.MousePointer = vbHourglass
    
    If bDevMode = False Then
        HASAR1.ReporteX
    Else
        Debug.Print "Reporte X."
    End If
    
    
    Screen.MousePointer = vbNormal
    
    Exit Sub

impresora_apag:

    Screen.MousePointer = vbNormal
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub


Private Sub mnuReporteZ_Click()
    
    On Error GoTo impresora_apag
    
Procesar:
    
    Screen.MousePointer = vbHourglass
    
    If bDevMode = False Then
        HASAR1.ReporteZ
    Else
        Debug.Print "Reporte X."
    End If
  
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
impresora_apag:

    Screen.MousePointer = vbNormal
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

' --------------------------------------------------------------------------
Private Sub mnuSysDebug_Click()

    mnuSysDebug.Checked = Not mnuSysDebug.Checked

    If mnuSysDebug.Checked Then
        
    Else
        MinimizeForm
    End If


End Sub


Private Function MaximizeForm()

    Me.Width = 11085

End Function

Private Function MinimizeForm()

    Me.Width = 4695

End Function

' --------------------------------------------------------------------------
'
' --------------------------------------------------------------------------
Private Function manageOrders(ByVal sData As String) As String

Dim iOrderCnt As Integer

'On Error GoTo ErrHandler

    ' esta imprimiendo?
    If bPrinting = False Then

        bPrinting = True
        ' veamos si hay algo para imprimir

        MuestraEstado "Consultando..."


cmd_GetPending:
        
        ' obtenemos las ordenes pendientes
        iOrderCnt = oTekBiz.GetPendingOrders(sData)
        
        ' marca las ordenes que procesamos como "emitidas"
        If oTekBiz.Error Then
            
            ' si hubo un error, preguntar si reintentamos
            Screen.MousePointer = vbNormal
            If MsgBox("Error: [" & oTekBiz.Error & "] " & oTekBiz.ErrDescription, vbRetryCancel, "Errores") = vbRetry Then
                GoTo cmd_GetPending
            End If
            
        End If
        
        '----------------------DEBUG
        If mnuSysDebug.Checked Then
            MuestraEstado "Error"
        End If
        '----------------------DEBUG

        If oTekBiz.Error Then
            ErrorHandler oTekBiz.ErrDescription
        End If
        
        
        ' hay algo para imprimrr=?
        If iOrderCnt > 0 Then
        
            MuestraEstado "Imprimiendo: " & iOrderCnt & " orden(es)."
        
            ' imprime los tickets
            ImprimirTickets oTekBiz
            
            
            manageOrders = oTekBiz.GetPrintedOrdersXML()
            
            If bDevMode = True Then
                Debug.Print manageOrders
            End If
            
            
'cmd_SetPrinted:
'
'            ' marca las ordenes que procesamos como "emitidas"
'            If Not oTekBiz.SetPrintedOrders Then
'
'                ' si hubo un error, preguntar si reintentamos
'                Screen.MousePointer = vbNormal
'                If MsgBox("Error: [" & oTekBiz.Error & "] " & oTekBiz.ErrDescription, vbRetryCancel, "Errores") = vbRetry Then
'                    GoTo cmd_SetPrinted
'                End If
'
'            End If
'
'            ' revisamos si hubo error es en la respuesta, no en la comunicacion
'            If oTekBiz.Error Then
'                Screen.MousePointer = vbNormal
'                If MsgBox("Error: [" & oTekBiz.Error & "] " & oTekBiz.ErrDescription, vbRetryCancel, "Errores") = vbRetry Then
'                    GoTo cmd_SetPrinted
'                End If
'            End If

        
            ' muestra las ordenes procesadas
            If bDebugMode = True Or bDevMode = True Then
                Debug.Print oTekBiz.OrdersProcessed
            End If
            
            ' inicializa la clase
            oTekBiz.Restart
        
            MuestraEstado iOrderCnt & " orden(es) procesada(s)"
        
        Else
        
            MuestraEstado "No hay ordenes para procesar."
        
        End If
        bPrinting = False
        
    End If


DeadEnd:

    
    Exit Function

ErrHandler:
    ErrorHandler

End Function


Private Sub Init()

    Screen.MousePointer = vbHourglass

    ' flag de impresion
    bPrinting = False
    
    ' levanta configuracion
    If Dir$(App.Path & "\config.ini") = "" Then
        Err.Raise 1, "Init", "No se encontro el archivo de configuracion."
    Else
    
        bDevMode = CBool(ReadINI("General", "dev", "config.ini"))
        sPrinterPort = ReadINI("General", "puerto", "config.ini")
    End If

    ' varios
    LogToFile = True
    frmMain.Caption = "Facturador " & ReadINI("General", "nombre", "config.ini")
    'frmMain.Icon = ReadINI("General", "icono", "config.ini")


    ' inicializamos el objeto de administracion de ordenes
    Set oTekBiz = New clsTekBiz
    If Not oTekBiz.Initialize Then
        MsgBox "No se pudo conectar con tekBiz, por favor comuniquese con el administrador."
        Screen.MousePointer = vbNormal
        Exit Sub
    End If


    ' inicializa el SOCKET
    Winsock.Protocol = sckTCPProtocol
    Winsock.LocalPort = sPrinterPort
    Winsock.Listen


    FS = Chr$(28)

    ' ahora inicializamos la impresora

'On Error GoTo impresora_apag
Procesar:

    If bDevMode = False Then
        ' inicializa control
        HASAR1.Puerto = ReadINI("Impresor", "puerto", "config.ini")
        HASAR1.AutodetectarModelo
        HASAR1.AutodetectarControlador
        HASAR1.Comenzar
        HASAR1.PrecioBase = False
        HASAR1.TratarDeCancelarTodo
    Else
        Debug.Print "Inicializa impresora."
    End If
    
    
    ' estado
    MuestraEstado "Listo."
    
    Screen.MousePointer = vbNormal
    
    Exit Sub

impresora_apag:

    Screen.MousePointer = vbNormal
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
    MuestraEstado "Error de comunicacion."
    Exit Sub

ErrHandler:
    ErrorHandler
End Sub

' --------------------------------------------------------------------------
' ELEMENTOS DEL SYSTRAY
' --------------------------------------------------------------------------

'[Checking The mouse event]
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim cEvent As Single

    cEvent = x / Screen.TwipsPerPixelX
    
    Select Case cEvent
        Case MouseMove
            'Debug.Print "MouseMove"
        Case LeftUp
            'Debug.Print "Left Up"
            'PopupMenu mnuForm
        Case LeftDown
            'Debug.Print "LeftDown"
        Case LeftDbClick
            'Debug.Print "LeftDbClick"
            Call mnuShow_Click
        Case MiddleUp
            'Debug.Print "MiddleUp"
        Case MiddleDown
            'Debug.Print "MiddleDown"
        Case MiddleDbClick
            'Debug.Print "MiddleDbClick"
        Case RightUp
            'Debug.Print "RightUp"
            PopupMenu mnuForm
        Case RightDown
            'Debug.Print "RightDown"
        Case RightDbClick
            'Debug.Print "RightDbClick"
    End Select

End Sub

Private Sub mnuHide_Click()
    If Not Me.WindowState = 1 Then WindowState = 1: Me.Hide
End Sub

Private Sub mnuShow_Click()
    If Me.WindowState = 1 Then WindowState = 0: Me.Show
    TrayDelete  '[Deleting Tray]
End Sub

' --------------------------------------------------------------------------
' EVENTOS DEL CONTROLADOR
' --------------------------------------------------------------------------
Private Sub HASAR1_ErrorFiscal(ByVal Flags As Long)
    
    If bDevMode = False Then
        If bDebugMode Then
            Debug.Print HASAR1.DescripcionStatusFiscal(Flags)
        End If
        ErrorLog HASAR1.DescripcionStatusFiscal(Flags)
        MuestraEstado HASAR1.DescripcionStatusFiscal(Flags)
    Else
        Debug.Print "Evento de error."
    End If

End Sub

Private Sub HASAR1_EventoFiscal(ByVal Flags As Long)
    
    If bDevMode = False Then
        If bDebugMode Then
            Debug.Print HASAR1.DescripcionStatusFiscal(Flags)
        End If
        '
        'ErrorLog HASAR1.DescripcionStatusFiscal(Flags)
        MuestraEstado HASAR1.DescripcionStatusFiscal(Flags)
    Else
        Debug.Print "Evento fiscal."
    End If
    
End Sub

Private Sub HASAR1_EventoImpresora(ByVal Flags As Long)

    If bDevMode = False Then
        If bDebugMode Then
            Debug.Print HASAR1.DescripcionStatusImpresor(Flags)
        End If
        
        ErrorLog HASAR1.DescripcionStatusImpresor(Flags)
    Else
        Debug.Print "Evento impresora."
    End If
    
    
    Select Case Flags
        Case P_JOURNAL_PAPER_LOW, P_RECEIPT_PAPER_LOW:
            If bDebugMode Then
                Debug.Print "Falta papel"
            End If
            ErrorLog "Falta papel"
        Case P_OFFLINE:
            If bDebugMode Then
                Debug.Print "Impresor fuera de l�nea"
            End If
            ErrorLog "Impresor fuera de l�nea"
        Case P_PRINTER_ERROR:
            If bDebugMode Then
                Debug.Print "Error mec�nico de impresor"
            End If
            ErrorLog "Error mec�nico de impresor"
        Case Else:
            If bDebugMode Then
                Debug.Print "Otro bit de impresora"
            End If
            ErrorLog "Otro bit de impresora"
    End Select

End Sub

Private Sub HASAR1_ImpresoraOcupada()
    If bDebugMode Then
        Debug.Print "DC2......."
    End If
End Sub


' --------------------------------------------------------------------------
'
' --------------------------------------------------------------------------
Sub TerminateApp()

    oTekBiz.Terminate
    Set oTekBiz = Nothing
    
    If bDevMode = False Then
        HASAR1.Finalizar
    Else
        Debug.Print "Cierra impresora."
    End If
    
    End

End Sub


Private Function GenerateDutyNumber() As String

    Dim sNumber As String, iNumber As Integer
    Dim aNumbers(1 To 3) As String
    
    aNumbers(1) = "354852M"
    aNumbers(2) = "958285R"
    aNumbers(3) = "915464V"

    Randomize
    
    ' generamos el numero
'    For i = 1 To 6
'        sNumber = sNumber & CStr(Int((9 - 1 + 1) * Rnd + 1))
'    Next

    ' le agregamos la letra final
    'sNumber = sNumber & Chr(Int((3 - 1 + 1) * Rnd + 1))
    iNumber = Int((3 - 1 + 1) * Rnd + 1)

    GenerateDutyNumber = aNumbers(iNumber)

End Function


Private Sub mnuSysReprint_Click()
    HASAR1.ReimprimirComprobante
End Sub

Private Sub Winsock_Close()
    
    Winsock.Close
    Winsock.Listen

End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)

    If Winsock.State = sckListening Then
        Winsock.Close
        Winsock.Accept requestID
    End If

End Sub

'Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
'    Dim sData As String
'    Dim sResponse As String
'
'    Winsock.GetData sData
'    DoEvents
'
'    ' Imprime comprobantes
'    If Winsock.State = sckConnected Then
'        If bDevMode = True Then
'            Debug.Print sData
'        End If
'
'        ' imprime
'        sResponse = manageOrders(sData)
'
'    End If
'
'    ' Envia respuesta
'    If Winsock.State = sckConnected Then
'        If Len(sResponse) = 0 Then
'            MsgBox "Error"
'            Winsock.SendData "ERROR"
'        Else
'            Winsock.SendData sResponse
'        End If
'    End If
'
'
'    DoEvents
'
'    Winsock_Close
'
'End Sub

'
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String, strLastPiece As String
    Dim lonEOP As Long, bolTruncated As Boolean
    Dim strPackets() As String, lonLoop As Long
    Dim lonUB As Long
    Dim sResponse As String
    
    'Get received data and put it into strData.
    Winsock.GetData strData, vbString, bytesTotal
    
    'Append it to end of buffer.
    strBuffer = strBuffer & strData
    
    'Erase string we just used.
    strData = vbNullString
    
    'See if there is a truncated packet.
    If Not Right$(strBuffer, 3) = EOP Then
        bolTruncated = True
        'Last packet received is incomplete and got cut off.
        'Remove this last piece and store it away for now.
        'Once we process the complete packets, we will put this back into the buffer
        'to be processed next time.
        
        'First, find last occurence of EOP.
        lonEOP = InStrRev(strBuffer, EOP)
        
        'If not found, then we haven't even received one full packet yet
        'so just exit.
        If lonEOP = 0 Then Exit Sub
        
        strLastPiece = Mid$(strBuffer, lonEOP + 2)
    End If
    
    'Done with that, now split up the data into individual packets.
    strPackets() = Split(strBuffer, EOP)
    lonUB = UBound(strPackets()) 'Number of items in array (# of packets).
    
    'Start looping through all packets.
    For lonLoop = 0 To lonUB
        If Len(strPackets(lonLoop)) > 3 Then 'A command is 3 bytes long.
            If bolTruncated And lonLoop = lonUB Then
                'This is the truncated one, do what you want with it here.
                'strPackets(lonLoop).
            Else
            
                'Mandar a imprimir
                sResponse = manageOrders(strPackets(lonLoop))
                
                ' Envia respuesta
                If Winsock.State = sckConnected Then
                    If Len(sResponse) = 0 Then
                        MsgBox "Error"
                        Winsock.SendData "ERROR"
                    Else
                        Winsock.SendData sResponse
                    End If
                End If
                
                DoEvents
                
                strBuffer = vbNullString
                
            
                Winsock_Close

            End If
        End If
    Next lonLoop
    
    'Clean up.
    Erase strPackets()
    
    If bolTruncated Then
        strBuffer = strLastPiece
    End If
    
End Sub



Public Sub MuestraEstado(ByVal texto As String)

    txtStatus.Text = texto & " | " & Time()

End Sub

Private Sub Winsock_SendComplete()
    
    Debug.Print "Listo"
    
End Sub
