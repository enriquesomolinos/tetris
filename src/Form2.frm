VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Mejores marcas"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox label3 
      Height          =   405
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox label2 
      Height          =   405
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox label1 
      Height          =   405
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Me.Hide

End Sub

Private Sub Form_Activate()
Dim prueba As record
Dim prueba3 As record

Dim prueba2 As record
If final = 1 Then GoTo t Else GoTo tt
t: Open "record.txt" For Random Access Write As #1 Len = Len(prueba)

label1 = prueba.nombre
Label4 = prueba.punt
label2 = prueba2.nombre
Label5 = prueba2.punt
label3 = prueba3.nombre
Label6 = prueba3.punt


Put #1, 1, prueba
Put #1, 2, prueba2

Put #1, 3, prueba

Close #1
tt:
label1.Enabled = False
label2.Enabled = False
label3.Enabled = False

Open "record.txt" For Random Access Read As #1 Len = Len(prueba)



Get #1, 1, prueba


Get #1, 2, prueba2




Get #1, 2, prueba3

'If fin = 1 Then GoTo a Else GoTo aaa
'a: If puntuacion > prueba.punt Then prueba3.punt = prueba2.punt: prueba3.nombre = prueba2.nombre: prueba2.punt = prueba.punt: prueba2.nombre = prueba.nombre: prueba.punt = puntuacion: prueba.nombre = "": label1.Enabled = True: label1.SetFocus: GoTo aaa
'If puntuacion > prueba2.punt Then prueba3.punt = prueba2.punt: prueba3.nombre = prueba2.nombre: prueba2.punt = prueba.punt: prueba2.punt = puntuacion: prueba2.nombre = "": label2.Enabled = True: label2.SetFocus: GoTo aaa
'If puntuacion > prueba3.punt Then prueba3.punt = puntuacion: prueba3.nombre = "": label3.Enabled = True: label3.SetFocus: GoTo aaa


aaa:
label1 = prueba.nombre
Label4 = prueba.punt
label2 = prueba2.nombre
Label5 = prueba2.punt
label3 = prueba3.nombre
Label6 = prueba3.punt
fin:
Close #1

End Sub

Private Sub Form_Load()
'Dim prueba As record
'prueba.punt = 8000
'prueba.nombre = "Enrique Somolinos"
'Open "record.txt" For Random Access Write As #1 Len = Len(prueba)

'open "record.txt" For Random Access Read As #1 Len = Len(prueba)



'Get #1, 1, prueba

'Put #1, 1, prueba
'Dim prueba2 As record
'Get #1, 2, prueba2



'prueba2.punt = 7100
'prueba2.nombre = "Enrique Somolinos"
'Put #1, 2, prueba2
'Dim prueba3 As record
'Get #1, 2, prueba3

'prueba3.punt = 6500
'prueba3.nombre = "Enrique Somolinos"
'Put #1, 3, prueba3
'Label1 = prueba.nombre
'Label4 = prueba.punt
'Label2 = prueba2.nombre
'Label5 = prueba2.punt
'label3 = prueba3.nombre
'Label6 = prueba3.punt
'Close #1
'Me.Hide
'Form1.Show

End Sub
