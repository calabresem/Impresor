VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{39ABE45D-F077-4D34-A361-6906C77D67F7}#1.0#0"; "Fiscal150423.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturador"
   ClientHeight    =   2055
   ClientLeft      =   2985
   ClientTop       =   3360
   ClientWidth     =   4620
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4620
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
      Left            =   480
      OleObjectBlob   =   "Form1.frx":1242
      Top             =   2280
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

Private Sub ImprimirTickets(ByRef oTekBiz As clsTekBiz)
    
'----------------------------------
' Codigos para items generales
'HASAR1.DescuentoUltimoItem "Oferta del Dia", 5, True
'HASAR1.DescuentoGeneral "Oferta Pago Efectivo", 25, True
'HASAR1.EspecificarPercepcionPorIVA "Percep IVA21", 100, 21
'HASAR1.EspecificarPercepcionGlobal "Percep. RG 0000", 125#


Dim x As Integer, Y As Integer
Dim sComando As String, sRazonSocial As String, sDomicilio As String, sCUIT As String
Dim sEmisorFMT As String, sNumeroFMT As String
Dim cTotDesc As Double
      
    ' hay algo para imprimir? (se supone que si, pero...)
    If oTekBiz.Orders.Count = 0 Then
        Exit Sub
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
                sRazonSocial = Left(oTekBiz.Orders.Item(x).RazonSocial, 30)
                sDomicilio = Left(oTekBiz.Orders.Item(x).Direccion, 40)
                ' le sacamos los guiones al CUIT
                sCUIT = Replace(oTekBiz.Orders.Item(x).CUIT, "-", "")
                
                ' completamos con espacios
                sRazonSocial = sRazonSocial & String((30 - Len(sRazonSocial)), " ")
                'sDomicilio = sDomicilio & String((40 - Len(sDomicilio)), " ")
                sDomicilio = String((40 - Len("")), " ")
                
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
                        Debug.Print "Imprime el detalle."
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

                Else
                    Debug.Print "Pago:  " & CStr(oTekBiz.Orders.Item(x).PagoEfectivo)
                    
                End If
                


            ' ------------------------------ TICKET FACTURA B
            Case Ticket
            Case FacturaB

               
               If bDevMode = False Then
                    HASAR1.DatosCliente "CONSUMIDOR FINAL", "00000000", TIPO_DNI, CONSUMIDOR_FINAL, "."
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

                Else
                    Debug.Print "Imprime el pago y cierre"
                End If



            ' ------------------------------ NOTA CREDITO A
            Case NotaCA

                If bDevMode = False Then
    
                    '// ----------------------------------------------
                    sRazonSocial = Left(oTekBiz.Orders.Item(x).RazonSocial, 30)
                    sDomicilio = Left(oTekBiz.Orders.Item(x).Direccion, 40)
                    sCUIT = Replace(oTekBiz.Orders.Item(x).CUIT, "-", "")
                    sRazonSocial = sRazonSocial & String((30 - Len(sRazonSocial)), " ")
                    sDomicilio = sDomicilio & String((40 - Len(sDomicilio)), " ")
                    
                    'HASAR1.InformacionRemito(1) = "0000-00000000"
                    'HASAR1.DatosCliente sRazonSocial, sCUIT, TIPO_CUIT, RESPONSABLE_INSCRIPTO, "."
                    'HASAR1.AbrirDNFH NOTA_CREDITO_A
    
                    HASAR1.Enviar Chr$(147) & FS & "1" & FS & "0000-0000000"
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

                Else
                    Debug.Print "Imprime pago y cierra."
                End If



            ' ------------------------------ NOTA CREDITO B
            Case NotaCB

                If bDevMode = False Then

                    ' Metodos no soportados por el 715F
                    'HASAR1.InformacionRemito(1) = "0000-00000000"
                    'HASAR1.DatosCliente "CONSUMIDOR FINAL", "99999999995", TIPO_CUIT, CONSUMIDOR_FINAL, "."
                    'HASAR1.AbrirDNFH NOTA_CREDITO_B
                    
                    HASAR1.Enviar Chr$(147) & FS & "1" & FS & "0000-0000000"
                    comando = Chr$(98) & FS & "CONSUMIDOR FINAL" & FS & "99999999995" & FS & "C" & FS & "2" & _
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

                Else
                    Debug.Print "Imprime detalle."
                End If
                
                
        Case Else
            Debug.Print "El comprobante no corresponde."
            

       
       End Select
   
   
    Next x
    
    ' puntero
    Screen.MousePointer = vbNormal

    Exit Sub

impresora_apag:

    Screen.MousePointer = vbNormal
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

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
Private Function manageOrders(ByVal sData As String) As Boolean

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


    manageOrders = True

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
                Debug.Print "Impresor fuera de línea"
            End If
            ErrorLog "Impresor fuera de línea"
        Case P_PRINTER_ERROR:
            If bDebugMode Then
                Debug.Print "Error mecánico de impresor"
            End If
            ErrorLog "Error mecánico de impresor"
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

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String

    Winsock.GetData sData
    
    If Winsock.State = sckConnected Then
        If bDevMode = True Then
            Debug.Print sData
        End If
        Winsock.SendData "OK"

        ' imprime
        If manageOrders(sData) = False Then
            MsgBox "Error"
        End If


    End If
    
    Winsock_Close

End Sub

Public Sub MuestraEstado(ByVal texto As String)

    txtStatus.Text = texto & " | " & Time()

End Sub
