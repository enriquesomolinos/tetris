VERSION 5.00
Begin VB.Form acercade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Madridkiller II"
   ClientHeight    =   3525
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5160
   ClipControls    =   0   'False
   Icon            =   "acerca.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "acerca.frx":030A
   ScaleHeight     =   2433.018
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   4845.507
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "acerca.frx":5368
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   2625
      Width           =   1245
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Información..."
      Height          =   345
      Left            =   3840
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4732.821
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(c) Enrique Somolinos Pérez 2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   3000
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Madridkiller II"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4732.821
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Advertencia: ..."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   2640
      Width           =   1080
   End
End
Attribute VB_Name = "acercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opciones de seguridad de claves del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos principales de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en Null
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
 Me.Hide
 
End Sub

Private Sub Form_Load()
    Me.Caption = "Acerca de " & App.Title
    
    lblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Prueba a obtener del Registro la información del sistema sobre el nombre y la ruta del programa...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Prueba a obtener del Registro la información del sistema sobre la ruta del programa...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Comprueba la existencia de una versión conocida de un archivo de 32 bits
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - Imposible encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Imposible encontrar la entrada de Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de dato de una clave de Registro
    Dim tmpVal As String                                    ' Almacén temporal de una valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de la variable de la clave de Registro
    '------------------------------------------------------------
    ' Abre la clave de Registro en la raíz {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abre la clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error...
    
    tmpVal = String$(1024, 0)                               ' Asigna espacio para la variable
    KeyValSize = 1024                                       ' Marca el tamaño de la variable
    
    '------------------------------------------------------------
    ' Recupera valores de claves de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtiene o crea un valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agrega una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Se encontró Null, se extrae de la cadena
    Else                                                    ' WinNT no tiene una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se encontró Null, sólo se extrae la cadena
    End If
    '------------------------------------------------------------
    ' Determina el tipo de valor de la clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Busca tipos de datos...
    Case REG_SZ                                             ' Tipo de dato de la cadena de la clave de Registro
        KeyVal = tmpVal                                     ' Copia el valor de la cadena
    Case REG_DWORD                                          ' El tipo de dato de la cadena de la clave es Double Word
        For i = Len(tmpVal) To 1 Step -1                    ' Convierte cada byte
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Genera el valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convierte Double Word a String
    End Select
    
    GetKeyValue = True                                      ' Vuelve con éxito
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Restaurar después de que ocurra un error...
    KeyVal = ""                                             ' Establece el valor de retorno para una cadena vacía
    GetKeyValue = False                                     ' Devuelve un error
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
End Function

