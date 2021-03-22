VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9948
   ClientLeft      =   2412
   ClientTop       =   1212
   ClientWidth     =   6588
   _ExtentX        =   11621
   _ExtentY        =   17547
   _Version        =   393216
   Description     =   "Se usa para cambiar un tipo de control por otro"
   DisplayName     =   "Project Examiner"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Private mcbMenuCommandBar         As Office.CommandBarControl
Private mfrmMain                 As New frmMain
Public WithEvents MenuHandler As CommandBarEvents          'controlador de evento de barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1


Sub HidefrmMain()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmMain.Hide
   
End Sub

Sub ShowfrmMain()
  
    On Error Resume Next
    
    If mfrmMain Is Nothing Then
        Set mfrmMain = New frmMain
    End If
    
    Set mfrmMain.VBInstance = VBInstance
    Set mfrmMain.Connect = Me
    FormDisplayed = True
    mfrmMain.Show
    mfrmMain.ZOrder
    mfrmMain.SetFocus
   
End Sub

'------------------------------------------------------
'este método agrega el complemento a VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'guardar la instanacia de vb
    Set VBInstance = Application
    
    'éste es un buen lugar para establecer un punto de interrupción y
    'y probar varios objetos, propiedades y métodos de complemento
    'Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Utilizado por la barra de herramientas de asistente para iniciar este asistente
        ShowfrmMain
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar(App.Title)
        'recibir el evento
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'establecer esto para mostrar el formulario al conectar
            ShowfrmMain
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'este método quita el complemento de VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'eliminar la entrada de la barra de comandos
    mcbMenuCommandBar.Delete
    
    'cerrar el complemento
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmMain
    Set mfrmMain = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'establecer esto para mostrar el formulario al conectar
        ShowfrmMain
    End If
End Sub

'este evento se desencadena cuando se hace clic en el menú desde el IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ShowfrmMain
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'objeto de barra de comandos
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'ver si podemos encontrar el menú Complementos
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'no disponible; error
        Exit Function
    End If
    
    'agregarlo a la barra de comandos
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'establecer el título
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

