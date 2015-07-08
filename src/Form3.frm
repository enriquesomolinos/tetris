VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   LinkTopic       =   "Form3"
   ScaleHeight     =   2700
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   548
      TabIndex        =   2
      Top             =   1883
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escribe tu nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enhorabuena tienes un récord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   788
      TabIndex        =   0
      Top             =   443
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Text1.Text = ""
Text1.MaxLength = 20
Text1.SetFocus
nomb = Text1.Text




End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then GoTo a Else GoTo fin
a:
Dim prueba As record
Dim prueba3 As record
Dim prueba2 As record
Open "record.txt" For Random Access Read As #1 Len = Len(prueba)
Get #1, 1, prueba
Get #1, 2, prueba2
Get #1, 2, prueba3

If puntuacion > prueba.punt Then prueba3.punt = prueba2.punt: prueba3.nombre = prueba2.nombre: prueba2.punt = prueba.punt: prueba2.nombre = prueba.nombre: prueba.punt = puntuacion: prueba.nombre = nomb: GoTo aaa
If puntuacion > prueba2.punt Then prueba3.punt = prueba2.punt: prueba3.nombre = prueba2.nombre: prueba2.punt = prueba.punt: prueba2.punt = puntuacion: prueba2.nombre = nomb: GoTo aaa
If puntuacion > prueba3.punt Then prueba3.punt = puntuacion: prueba3.nombre = nomb:  GoTo aaa



aaa:
Close #1

Me.Hide
Form2.Show (1)
fin:

End Sub
