VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris por Enrique Somolinos"
   ClientHeight    =   5325
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4425
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000A&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   720
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   136
      Top             =   2760
      Width           =   615
   End
   Begin VB.PictureBox pieza7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3600
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   720
      TabIndex        =   135
      Top             =   4320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pieza6 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1440
      Picture         =   "Form1.frx":154C
      ScaleHeight     =   480
      ScaleWidth      =   705
      TabIndex        =   134
      Top             =   4800
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox pieza5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2160
      Picture         =   "Form1.frx":278E
      ScaleHeight     =   480
      ScaleWidth      =   720
      TabIndex        =   133
      Top             =   4800
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pieza3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3120
      Picture         =   "Form1.frx":39D0
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   132
      Top             =   4920
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox pieza4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2400
      Picture         =   "Form1.frx":4B82
      ScaleHeight     =   480
      ScaleWidth      =   705
      TabIndex        =   131
      Top             =   4800
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Pieza2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1680
      Picture         =   "Form1.frx":5DC4
      ScaleHeight     =   240
      ScaleWidth      =   960
      TabIndex        =   130
      Top             =   4560
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Pieza1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3840
      Picture         =   "Form1.frx":6A06
      ScaleHeight     =   480
      ScaleWidth      =   465
      TabIndex        =   129
      Top             =   4920
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   128
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   128
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   127
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   127
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   126
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   126
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   125
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   125
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   124
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   124
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   123
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   123
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   122
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   122
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   121
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   121
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   120
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   120
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   119
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   119
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   118
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   118
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   117
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   117
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   116
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   116
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   115
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   115
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   114
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   114
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   113
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   113
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   112
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   112
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   111
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   111
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   110
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   110
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   109
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   109
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   108
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   108
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   107
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   107
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   106
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   106
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   105
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   105
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   104
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   104
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   103
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   103
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   102
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   102
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   101
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   101
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   100
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   100
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   99
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   99
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   98
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   98
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   97
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   97
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   96
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   96
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   95
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   95
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   94
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   94
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   93
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   93
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   92
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   92
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   91
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   91
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   90
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   90
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   89
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   89
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   88
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   88
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   87
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   87
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   86
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   86
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   85
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   85
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   84
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   84
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   83
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   83
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   82
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   82
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   81
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   81
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   80
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   80
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   79
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   79
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   78
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   78
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   77
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   77
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   76
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   76
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   75
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   75
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   74
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   74
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   73
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   73
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   72
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   72
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   71
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   71
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   70
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   70
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   69
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   69
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   68
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   68
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   67
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   67
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   66
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   66
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   65
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   65
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   64
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   64
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   63
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   63
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   62
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   62
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   61
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   61
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   60
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   60
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   59
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   59
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   58
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   58
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   57
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   57
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   56
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   55
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   54
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   53
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   53
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   52
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   51
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   51
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   50
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   50
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   49
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   49
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   48
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   48
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   47
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   47
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   46
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   46
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   45
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   45
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   44
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   44
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   43
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   43
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   42
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   42
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   41
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   41
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   40
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   40
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   39
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   39
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   38
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   38
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   37
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   37
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   36
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   36
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   35
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   35
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   34
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   34
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   33
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   33
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   32
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   32
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   31
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   31
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   30
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   30
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   29
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   29
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   28
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   28
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   27
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   27
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   26
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   26
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   25
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   25
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   24
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   24
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   23
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   22
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   22
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   21
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   21
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   20
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   19
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   18
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   17
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   16
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   15
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   14
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   13
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   12
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   11
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   10
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   9
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   8
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   7
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   6
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   5
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   4
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   3
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   2
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   1
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Tag             =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   0
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Perdiste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2640
      TabIndex        =   148
      Top             =   4320
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Siguiente pieza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   262
      TabIndex        =   147
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Puntuación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   146
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nivel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   765
      TabIndex        =   145
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Lineas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   144
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Siguiente récord:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   143
      Top             =   3840
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Récord total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   397
      TabIndex        =   142
      Top             =   4440
      Width           =   1350
   End
   Begin VB.Label Lpuntuacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   195
      Left            =   840
      TabIndex        =   141
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Lnivel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   195
      Left            =   832
      TabIndex        =   140
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Llineas 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   195
      Left            =   832
      TabIndex        =   139
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Lsig 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "6500"
      Height          =   195
      Left            =   900
      TabIndex        =   138
      Top             =   4200
      Width           =   360
   End
   Begin VB.Label Lrec 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "8000"
      Height          =   195
      Left            =   900
      TabIndex        =   137
      Top             =   4800
      Width           =   360
   End
   Begin VB.Menu Juego 
      Caption         =   "&Juego"
      Begin VB.Menu nueva 
         Caption         =   "Nueva partida"
      End
      Begin VB.Menu marcas 
         Caption         =   "Mejores marcas"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu acerca 
         Caption         =   "Acerca de Tetris"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acerca_Click()
acercade.Show
End Sub

Private Sub Form_Load()
Randomize
Dim prueba As record

Open "record.txt" For Random Access Read As #1 Len = Len(prueba)



Get #1, 1, prueba

Dim prueba2 As record
Get #1, 2, prueba2




Dim prueba3 As record
Get #1, 2, prueba3




Lsig = prueba3.punt
Lrec = prueba.punt

perder = 0
puntuacion = 0
nivel = 0
lineas = 0

Lpuntuacion = puntuacion
Lnivel = nivel
Llineas = lineas
lrecord = record
Close #1



x = 49
index = x
rot = 1: caer = 0

pieza = Int(Rnd * 7) + 1

'cubo
If pieza = 1 Then Picture1(x).BackColor = &H80C0FF: Picture1(x + 1).BackColor = &H80C0FF: Picture1(x + 16).BackColor = &H80C0FF: Picture1(x + 17).BackColor = &H80C0FF: Picture2.Picture = Pieza1.Picture
If pieza = 2 Then x = 65: Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 2).BackColor = &H8000000F: Picture2.Picture = Pieza2.Picture
If pieza = 3 Then x = 49: Picture1(x).BackColor = &HFF00&: Picture1(x + 16).BackColor = &HFF00&: Picture1(x + 1).BackColor = &HFF00&: Picture1(x - 15).BackColor = &HFF00&: Picture2.Picture = pieza3.Picture
If pieza = 4 Then x = 49: Picture1(x).BackColor = &HFF00FF: Picture1(x + 17).BackColor = &HFF00FF: Picture1(x + 1).BackColor = &HFF00FF: Picture1(x - 16).BackColor = &HFF00FF: Picture2.Picture = pieza4.Picture
If pieza = 5 Then x = 49: Picture1(x).BackColor = &HFF0000: Picture1(x + 16).BackColor = &HFF0000: Picture1(x - 16).BackColor = &HFF0000: Picture1(x - 15).BackColor = &HFF0000: Picture2.Picture = pieza5.Picture
If pieza = 6 Then x = 49: Picture1(x).BackColor = &H404080: Picture1(x + 16).BackColor = &H404080: Picture1(x - 16).BackColor = &H404080: Picture1(x + 17).BackColor = &H404080: Picture2.Picture = pieza6.Picture
If pieza = 7 Then x = 49: Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0:: Picture2.Picture = pieza7.Picture
piezasig = Int(Rnd(1) * 7) + 1
If piezasig = 1 Then Picture2.Picture = Pieza1.Picture
If piezasig = 2 Then Picture2.Picture = Pieza2.Picture
If piezasig = 3 Then Picture2.Picture = pieza3.Picture
If piezasig = 4 Then Picture2.Picture = pieza4.Picture
If piezasig = 5 Then Picture2.Picture = pieza5.Picture
If piezasig = 6 Then Picture2.Picture = pieza6.Picture
If piezasig = 7 Then Picture2.Picture = pieza7.Picture

End Sub





Private Sub marcas_Click()
Form2.Show (1)

End Sub

Private Sub nueva_Click()
For a = 1 To 113 Step 16
For h = 0 To 15
Picture1(a + h).BackColor = Picture1(0).BackColor

Next h
Next a
Label50.Visible = False
fin = 0
caer = 0


perder = 0
puntuacion = 0
nivel = 0
lineas = 0

Lpuntuacion = puntuacion
Lnivel = nivel
Llineas = lineas
lrecord = record



x = 49
index = x
rot = 1: caer = 0

pieza = Int(Rnd * 7) + 1

'cubo
If pieza = 1 Then Picture1(x).BackColor = &H80C0FF: Picture1(x + 1).BackColor = &H80C0FF: Picture1(x + 16).BackColor = &H80C0FF: Picture1(x + 17).BackColor = &H80C0FF: Picture2.Picture = Pieza1.Picture
If pieza = 2 Then x = 65: Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 2).BackColor = &H8000000F: Picture2.Picture = Pieza2.Picture
If pieza = 3 Then x = 49: Picture1(x).BackColor = &HFF00&: Picture1(x + 16).BackColor = &HFF00&: Picture1(x + 1).BackColor = &HFF00&: Picture1(x - 15).BackColor = &HFF00&: Picture2.Picture = pieza3.Picture
If pieza = 4 Then x = 49: Picture1(x).BackColor = &HFF00FF: Picture1(x + 17).BackColor = &HFF00FF: Picture1(x + 1).BackColor = &HFF00FF: Picture1(x - 16).BackColor = &HFF00FF: Picture2.Picture = pieza4.Picture
If pieza = 5 Then x = 49: Picture1(x).BackColor = &HFF0000: Picture1(x + 16).BackColor = &HFF0000: Picture1(x - 16).BackColor = &HFF0000: Picture1(x - 15).BackColor = &HFF0000: Picture2.Picture = pieza5.Picture
If pieza = 6 Then x = 49: Picture1(x).BackColor = &H404080: Picture1(x + 16).BackColor = &H404080: Picture1(x - 16).BackColor = &H404080: Picture1(x + 17).BackColor = &H404080: Picture2.Picture = pieza6.Picture
If pieza = 7 Then x = 49: Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0:: Picture2.Picture = pieza7.Picture
piezasig = Int(Rnd(1) * 7) + 1
If piezasig = 1 Then Picture2.Picture = Pieza1.Picture
If piezasig = 2 Then Picture2.Picture = Pieza2.Picture
If piezasig = 3 Then Picture2.Picture = pieza3.Picture
If piezasig = 4 Then Picture2.Picture = pieza4.Picture
If piezasig = 5 Then Picture2.Picture = pieza5.Picture
If piezasig = 6 Then Picture2.Picture = pieza6.Picture
If piezasig = 7 Then Picture2.Picture = pieza7.Picture
Close #3
End Sub

Private Sub Picture1_KeyPress(index As Integer, KeyAscii As Integer)
If fin = 1 Then GoTo fin
If KeyAscii = 52 Then GoSub izq
If KeyAscii = 54 Then GoSub der
If KeyAscii = 53 Then GoSub bajar
If KeyAscii = 56 Then GoSub rotar
GoTo fin

rotar:
        If pieza = 2 Then GoTo rot2 Else GoTo rot3
rot2: If rot = 1 Then GoTo 20 Else GoTo rot22:
20: For a = 1 To 113 Step 16
    If x = a Then Return
    Next
    For b = 16 To 128 Step 16
    If x = b Then Return
    Next
    For c = 15 To 127 Step 16
    If x = c Then Return
    Next

If Picture1(x + 14).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F And Picture1(x - 31).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



rot22: If rot = 2 Then GoTo rot21 Else GoTo rot3

rot21: For a = 1 To 16
    If x - 16 = a Or x = a Then Return
    Next
        For b = 113 To 128
        If x = b Then Return
        Next
        
If Picture1(x + 15).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x - 31).BackColor = &H8000000F And Picture1(x - 16).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F And Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 13).BackColor = &H8000000F Then GoTo borrar Else GoTo fin




rot3: If pieza = 3 Then GoTo rot3verde Else GoTo rot4
rot3verde: If rot = 1 Then GoTo rot31 Else GoTo rot32
rot31: For a = 16 To 128 Step 16
        If x = a Then Return
        Next
        For s = 1 To 113 Step 16
        If x = s Then Return
        Next
        If Picture1(x - 1).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F And Picture1(x + 2).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


rot32: If rot = 2 Then GoTo 58 Else GoTo rot4
58: For a = 1 To 16 Step 1
    If x = a Then Return
    Next
         If Picture1(x + 15).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



rot4: If pieza = 4 Then GoTo rot444 Else GoTo rot5

rot444: If rot = 1 Then GoTo rot41 Else If rot = 2 Then GoTo rot42
rot41: For a = 16 To 128 Step 16
        If x = a Then Return
        Next
        For s = 1 To 113 Step 16
        If x = s Then Return
        Next
        If Picture1(x - 1).BackColor = &H8000000F And Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


rot42: If rot = 2 Then GoTo rot422
rot422: For a = 113 To 128 Step 1
    If x = a Then Return
    Next
         If Picture1(x - 17).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin





rot5: If pieza = 5 Then GoTo rot55 Else GoTo rot6
rot55: If rot = 1 Then GoTo rot51 Else GoTo rot52
rot51: For a = 1 To 113 Step 16
        If x = a Then Return
        Next
        
        If Picture1(x - 1).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

rot52: If rot = 2 Then GoTo rot522 Else GoTo rot53
rot522: For a = 1 To 16 Step 1
        If x = a Then Return
        Next
        
        If Picture1(x - 16).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



rot53: If rot = 3 Then GoTo rot533 Else GoTo rot54
rot533: For a = 16 To 128 Step 16
        If x = a Then Return
        Next
        
        If Picture1(x - 1).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


rot54: If rot = 4 Then GoTo rot544 Else GoTo rot6
rot544: For a = 113 To 128 Step 1
        If x = a Then Return
        Next
       
        If Picture1(x - 16).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin




rot6: If pieza = 6 Then GoTo rot61 Else GoTo rot7
rot61: If rot = 1 Then GoTo rot611 Else GoTo rot62
rot611: For a = 1 To 113 Step 16
        If x = a Then Return
        Next
        
        If Picture1(x - 1).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


rot62: If rot = 2 Then GoTo rot622 Else GoTo rot63
rot622:
For a = 1 To 16 Step 1
        If x = a Then Return
        Next
       
        If Picture1(x - 16).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

rot63: If rot = 3 Then GoTo rot633 Else GoTo rot64
rot633: For a = 16 To 128 Step 16
        If x = a Then Return
        Next
        
        If Picture1(x - 1).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

rot64: If rot = 4 Then GoTo rot644 Else GoTo rot7
rot644: For a = 113 To 128 Step 1
        If x = a Then Return
        Next
        
        If Picture1(x - 16).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

rot7: If pieza = 7 Then GoTo rot71 Else Return
rot71: If rot = 1 Then GoTo rot711 Else GoTo rot72
rot711: For a = 1 To 113 Step 16
        If x = a Then Return
        Next
        
        If Picture1(x - 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


rot72: If rot = 2 Then GoTo rot722 Else GoTo rot73
rot722:
For a = 1 To 16 Step 1
        If x = a Then Return
        Next
       
        If Picture1(x - 16).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

rot73: If rot = 3 Then GoTo rot733 Else GoTo rot74
rot733: For a = 16 To 128 Step 16
        If x = a Then Return
        Next
        
        If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

rot74: If rot = 4 Then GoTo rot744 Else Return
rot744: For a = 113 To 128 Step 1
        If x = a Then Return
        Next
        
        If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

Return







bajar: If pieza = 1 Then GoTo cubobajar Else GoTo palob
cubobajar: For c = 1 To 8 Step 1
            If x + 1 = c * 16 Then caer = 1: Return
 Next c
        If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



palob: If pieza = 2 Then GoTo palob2 Else GoTo pieza3b
palob2:  If rot = 2 Then GoTo palob2rot1 Else GoTo palob2rot2
palob2rot1:            For d = 1 To 8 Step 1
            If x + 2 = d * 16 Then caer = 1: Return
Next d
            If Picture1(x + 3).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


palob2rot2: If rot = 1 Then GoTo 30 Else GoTo pieza3b
30: For d = 1 To 8 Step 1
            If x = d * 16 Then caer = 1: Return
Next d
            If Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 31).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin




pieza3b: If pieza = 3 Then GoTo p3baj Else GoTo p4baj
p3baj: If rot = 1 Then GoTo rot1pi3 Else GoTo rot2pi3
rot1pi3: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

rot2pi3: For b = 1 To 8 Step 1
        If x + 17 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 18).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


p4baj: If pieza = 4 Then GoTo pi4baj Else GoTo pi5baj
pi4baj: If rot = 1 Then GoTo pi41baj Else If rot = 2 Then GoTo pi42baj
pi41baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi42baj: If rot = 2 Then GoTo pi422baj Else GoTo pi5baj
pi422baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoTo borrar Else GoTo fin





pi5baj: If pieza = 5 Then GoTo pi511baj Else GoTo pi6baj
pi511baj: If rot = 1 Then GoTo pi51baj Else GoTo pi52baj
pi51baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi52baj: If rot = 2 Then GoTo pi522baj Else GoTo pi53baj
pi522baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi53baj: If rot = 3 Then GoTo pi533baj Else GoTo pi54baj
pi533baj: For b = 1 To 8 Step 1
        If x = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



pi54baj: If rot = 4 Then GoTo pi544baj Else GoTo pi6baj

pi544baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 16).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi6baj: If pieza = 6 Then GoTo pi611baj Else GoTo pi7baj
pi611baj: If rot = 1 Then GoTo pi61baj Else GoTo pi62baj
pi61baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi62baj: If rot = 2 Then GoTo pi622baj Else GoTo pi63baj
pi622baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F Then GoTo borrar Else GoTo fin





pi63baj: If rot = 3 Then GoTo pi633baj Else GoTo pi64baj
pi633baj: For b = 1 To 8 Step 1
        If x = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



pi64baj: If rot = 4 Then GoTo pi644baj Else GoTo pi7baj

pi644baj:  For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoTo borrar Else GoTo fin






pi7baj: If pieza = 7 Then GoTo pi711baj
pi711baj: If rot = 1 Then GoTo pi71baj Else GoTo pi72baj
pi71baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi72baj: If rot = 2 Then GoTo pi722baj Else GoTo pi73baj
pi722baj: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin





pi73baj: If rot = 3 Then GoTo pi733baj Else GoTo pi74baj
pi733baj: For b = 1 To 8 Step 1
        If x = b * 16 Then caer = 1: Return
        Next
        
          If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



pi74baj: If rot = 4 Then GoTo pi744baj

pi744baj:  For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: Return
        Next


          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin





izq:            If pieza = 1 Then GoTo cubo Else GoTo palo
cubo:                  For Y = 1 To 16 Step 1
            If x = Y Then Return:
            Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

palo:             If pieza = 2 Then GoTo palo2 Else GoTo pieza3
palo2:  If rot = 2 Then GoTo palo2rot1 Else GoTo palo2rot2
palo2rot1:      For r = 1 To 16 Step 1
                If x = r Then Return:
                Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

palo2rot2: If rot = 1 Then GoTo 40 Else GoTo pieza3:
40:             For r = 1 To 16 Step 1
                If x - 32 = r Then Return:
                Next: If Picture1(x - 48).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pieza3: If pieza = 3 Then GoTo izq3 Else GoTo izq4
izq3: If rot = 1 Then GoTo izq3r1 Else GoTo izq3r2
izq3r1: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 31).BackColor = &H8000000F And Picture1(x - 16).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


izq3r2: For r = 1 To 16 Step 1
                If x = r Then Return:
                Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

izq4: If pieza = 4 Then GoTo pi4izq Else GoTo pi5izq
pi4izq: If rot = 1 Then GoTo pi41izq Else If rot = 2 Then GoTo pi42izq
pi41izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi42izq: If rot = 2 Then GoTo pi422izq Else GoTo pi5izq
pi422izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 33).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi5izq: If pieza = 5 Then GoTo pi51izq Else GoTo pi6izq
pi51izq: If rot = 1 Then GoTo pi511izq Else GoTo pi52izq
pi511izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 31).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi52izq: If rot = 2 Then GoTo pi522izq Else GoTo pi53izq:
pi522izq: For r = 1 To 16 Step 1
                If x = r Then Return:
                Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi53izq: If rot = 3 Then GoTo pi533izq Else GoTo pi54izq
pi533izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



pi54izq: If rot = 4 Then GoTo pi544izq Else GoTo pi6izq
pi544izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x - 33).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi6izq: If pieza = 6 Then GoTo pi61izq Else GoTo pi7izq
pi61izq: If rot = 1 Then GoTo pi611izq Else GoTo pi62izq
pi611izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi62izq: If rot = 2 Then GoTo pi622izq Else GoTo pi63izq:
pi622izq: For r = 1 To 16 Step 1
                If x = r Then Return:
                Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi63izq: If rot = 3 Then GoTo pi633izq Else GoTo pi64izq
pi633izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 33).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



pi64izq: If rot = 4 Then GoTo pi644izq Else GoTo pi7izq
pi644izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 31).BackColor = &H8000000F Then GoTo borrar Else GoTo fin






pi7izq: If pieza = 7 Then GoTo pi71izq
pi71izq: If rot = 1 Then GoTo pi711izq Else GoTo pi72izq
pi711izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi72izq: If rot = 2 Then GoTo pi722izq Else GoTo pi73izq:
pi722izq: For r = 1 To 16 Step 1
                If x = r Then Return:
                Next: If Picture1(x - 16).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi73izq: If rot = 3 Then GoTo pi733izq Else GoTo pi74izq
pi733izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



pi74izq: If rot = 4 Then GoTo pi744izq
pi744izq: For r = 1 To 16 Step 1
                If x - 16 = r Then Return:
                Next: If Picture1(x - 32).BackColor = &H8000000F And Picture1(x - 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin





der: If pieza = 1 Then GoTo cubo2 Else GoTo palod
cubo2:         For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
                Next: If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 33).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


palod:          If pieza = 2 Then GoTo palod2 Else GoTo pieza3d
palod2:  If rot = 2 Then GoTo palod2rot1 Else GoTo palod2rot2
palod2rot1:            For b = 113 To 128 Step 1
                If x = b Then Return:
Next
 If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

palod2rot2: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 32).BackColor = &H8000000F Then GoTo borrar Else GoTo fin



pieza3d: If pieza = 3 Then GoTo pi3derr1 Else GoTo pi4derr1
pi3derr1: If rot = 1 Then GoTo tony Else GoTo pi3derr2

tony: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi3derr2: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 33).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin






pi4derr1: If pieza = 4 Then GoTo pi4der Else GoTo pi5der
pi4der: If rot = 1 Then GoTo pi41der Else If rot = 2 Then GoTo pi42der
pi41der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 33).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi42der: If rot = 2 Then GoTo pi422der Else GoTo pi5der
pi422der: For b = 113 To 128 Step 1
                If x = b Then Return:
Next
 If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi5der: If pieza = 5 Then GoTo pi51der Else GoTo pi6der
pi51der: If rot = 1 Then GoTo pi511der Else GoTo pi52der
pi511der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
            Next
            If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi52der: If rot = 2 Then GoTo pi522der Else GoTo pi53der
pi522der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F And Picture1(x + 33).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi53der: If rot = 3 Then GoTo pi533der Else GoTo pi54der
pi533der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 31).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi54der: If rot = 4 Then GoTo pi544der Else GoTo pi6der
pi544der: For b = 113 To 128 Step 1
                If x = b Then Return:
Next
 If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin





pi6der: If pieza = 6 Then GoTo pi61der Else GoTo pi7der
pi61der: If rot = 1 Then GoTo pi611der Else GoTo pi62der
pi611der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
            Next
            If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 33).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi62der: If rot = 2 Then GoTo pi622der Else GoTo pi63der
pi622der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 31).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi63der: If rot = 3 Then GoTo pi633der Else GoTo pi64der
pi633der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 32).BackColor = &H8000000F And Picture1(x - 1).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi64der: If rot = 4 Then GoTo pi644der Else GoTo pi7der
pi644der: For b = 113 To 128 Step 1
                If x = b Then Return:
Next
 If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin






pi7der: If pieza = 7 Then GoTo pi71der
pi71der: If rot = 1 Then GoTo pi711der Else GoTo pi72der
pi711der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
            Next
            If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi72der: If rot = 2 Then GoTo pi722der Else GoTo pi73der
pi722der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin


pi73der: If rot = 3 Then GoTo pi733der Else GoTo pi74der
pi733der: For b = 113 To 128 Step 1
                If x + 16 = b Then Return:
Next
 If Picture1(x + 32).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin

pi74der: If rot = 4 Then GoTo pi744der
pi744der: For b = 113 To 128 Step 1
                If x = b Then Return:
Next
 If Picture1(x + 16).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x + 15).BackColor = &H8000000F Then GoTo borrar Else GoTo fin







borrar:
If pieza = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F
If pieza = 2 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 2).BackColor = &H8000000F
If pieza = 2 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 32).BackColor = &H8000000F
If pieza = 3 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F
If pieza = 3 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F
If pieza = 4 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F
If pieza = 4 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F
If pieza = 5 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F
If pieza = 5 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F
If pieza = 5 And rot = 3 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 15).BackColor = &H8000000F
If pieza = 5 And rot = 4 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 17).BackColor = &H8000000F
If pieza = 6 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F
If pieza = 6 And rot = 4 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F
If pieza = 6 And rot = 3 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 17).BackColor = &H8000000F
If pieza = 6 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 15).BackColor = &H8000000F
If pieza = 7 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F
If pieza = 7 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F
If pieza = 7 And rot = 3 Then Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F
If pieza = 7 And rot = 4 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F



If KeyAscii = 52 Then x = x - 16
If KeyAscii = 54 Then x = x + 16
If KeyAscii = 53 Then x = x + 1
If KeyAscii = 56 Then rot = rot + 1
If pieza = 2 And rot = 3 Then rot = 1
If pieza = 3 And rot = 3 Then rot = 1
If pieza = 4 And rot = 3 Then rot = 1
If pieza = 5 And rot = 5 Then rot = 1
If pieza = 6 And rot = 5 Then rot = 1
If pieza = 7 And rot = 5 Then rot = 1




dibujar:
If pieza = 1 Then Picture1(x).BackColor = &H80C0FF: Picture1(x + 1).BackColor = &H80C0FF: Picture1(x + 16).BackColor = &H80C0FF: Picture1(x + 17).BackColor = &H80C0FF
If pieza = 2 And rot = 2 Then Picture1(x).BackColor = &HFF&: Picture1(x - 1).BackColor = &HFF&: Picture1(x + 1).BackColor = &HFF&: Picture1(x + 2).BackColor = &HFF&
If pieza = 2 And rot = 1 Then Picture1(x).BackColor = &HFF&: Picture1(x + 16).BackColor = &HFF&: Picture1(x - 16).BackColor = &HFF&: Picture1(x - 32).BackColor = &HFF&
If pieza = 3 And rot = 1 Then Picture1(x).BackColor = &HFF00&: Picture1(x + 16).BackColor = &HFF00&: Picture1(x + 1).BackColor = &HFF00&: Picture1(x - 15).BackColor = &HFF00&
If pieza = 3 And rot = 2 Then Picture1(x).BackColor = &HFF00&: Picture1(x + 17).BackColor = &HFF00&: Picture1(x + 16).BackColor = &HFF00&: Picture1(x - 1).BackColor = &HFF00&
If pieza = 4 And rot = 1 Then Picture1(x).BackColor = &HFF00FF: Picture1(x + 17).BackColor = &HFF00FF: Picture1(x + 1).BackColor = &HFF00FF: Picture1(x - 16).BackColor = &HFF00FF
If pieza = 4 And rot = 2 Then Picture1(x).BackColor = &HFF00FF: Picture1(x - 1).BackColor = &HFF00FF: Picture1(x - 16).BackColor = &HFF00FF: Picture1(x - 15).BackColor = &HFF00FF
If pieza = 5 And rot = 1 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 16).BackColor = &HFF0000: Picture1(x - 16).BackColor = &HFF0000: Picture1(x - 15).BackColor = &HFF0000
If pieza = 5 And rot = 2 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 17).BackColor = &HFF0000: Picture1(x - 1).BackColor = &HFF0000: Picture1(x + 1).BackColor = &HFF0000
If pieza = 5 And rot = 3 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 16).BackColor = &HFF0000: Picture1(x - 16).BackColor = &HFF0000: Picture1(x + 15).BackColor = &HFF0000
If pieza = 5 And rot = 4 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 1).BackColor = &HFF0000: Picture1(x - 1).BackColor = &HFF0000: Picture1(x - 17).BackColor = &HFF0000
If pieza = 6 And rot = 1 Then Picture1(x).BackColor = &H404080: Picture1(x + 16).BackColor = &H404080: Picture1(x - 16).BackColor = &H404080: Picture1(x + 17).BackColor = &H404080
If pieza = 6 And rot = 4 Then Picture1(x).BackColor = &H404080: Picture1(x + 1).BackColor = &H404080: Picture1(x - 1).BackColor = &H404080: Picture1(x - 15).BackColor = &H404080
If pieza = 6 And rot = 3 Then Picture1(x).BackColor = &H404080: Picture1(x + 16).BackColor = &H404080: Picture1(x - 16).BackColor = &H404080: Picture1(x - 17).BackColor = &H404080
If pieza = 6 And rot = 2 Then Picture1(x).BackColor = &H404080: Picture1(x + 1).BackColor = &H404080: Picture1(x - 1).BackColor = &H404080: Picture1(x + 15).BackColor = &H404080
If pieza = 7 And rot = 1 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0
If pieza = 7 And rot = 2 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0: Picture1(x - 1).BackColor = &H3DFAF0
If pieza = 7 And rot = 3 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x - 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0
If pieza = 7 And rot = 4 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x - 1).BackColor = &H3DFAF0



fin:

End Sub

Private Sub Salir_Click()
final = 1

Form2.Show
Me.Hide

1
End
End Sub


Private Sub Timer1_Timer()

If caer = 1 Then GoTo 1 Else GoTo 2
1: If x = 49 Or x = 65 Or x = 50 Or x = 66 Then fin = 1: Label50.Visible = True: Form3.Show (1): GoTo fin
2: If nivel = 0 Then Timer1.Interval = 1000
 If nivel = 1 Then Timer1.Interval = 900
 If nivel = 2 Then Timer1.Interval = 800
 If nivel = 3 Then Timer1.Interval = 700
 If nivel = 4 Then Timer1.Interval = 600
 If nivel = 5 Then Timer1.Interval = 500
 If nivel = 6 Then Timer1.Interval = 400
 If nivel = 7 Then Timer1.Interval = 300
 If nivel = 8 Then Timer1.Interval = 250
 If nivel = 9 Then Timer1.Interval = 200
 If nivel = 10 Then Timer1.Interval = 150
 If nivel = 11 Then Timer1.Interval = 100
 If nivel = 12 Then Timer1.Interval = 50

 
 
 
 
 If fin = 1 Then GoTo fin
  GoSub borrar:  GoSub comprobar: GoSub dibujar: GoTo fin


borrar: If caer = 1 Then GoSub genera: Return
If pieza = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Return
If pieza = 2 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 2).BackColor = &H8000000F: Return
If pieza = 2 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 32).BackColor = &H8000000F: Return
If pieza = 3 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F: Return
If pieza = 3 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Return
If pieza = 4 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Return
If pieza = 4 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F: Return
If pieza = 5 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F: Return
If pieza = 5 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Return
If pieza = 5 And rot = 3 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 15).BackColor = &H8000000F: Return
If pieza = 5 And rot = 4 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 17).BackColor = &H8000000F: Return
If pieza = 6 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 17).BackColor = &H8000000F: Return
If pieza = 6 And rot = 4 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 15).BackColor = &H8000000F: Return
If pieza = 6 And rot = 3 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 17).BackColor = &H8000000F: Return
If pieza = 6 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x + 15).BackColor = &H8000000F: Return
If pieza = 7 And rot = 1 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Return
If pieza = 7 And rot = 2 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Return
If pieza = 7 And rot = 3 Then Picture1(x).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x + 16).BackColor = &H8000000F: Return
If pieza = 7 And rot = 4 Then Picture1(x).BackColor = &H8000000F: Picture1(x + 1).BackColor = &H8000000F: Picture1(x - 16).BackColor = &H8000000F: Picture1(x - 1).BackColor = &H8000000F: Return







comprobar: If pieza = 1 Then GoTo cubo Else GoTo Pieza2
cubo: For a = 1 To 8 Step 1
            If x + 1 = a * 16 Then caer = 1: GoSub dibujar: GoSub genera
            
Next
      
    If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return
        
        
Pieza2:        If pieza = 2 Then GoTo palo Else GoTo pieza3
palo: If rot = 2 Then GoTo palorot1 Else If rot = 1 Then GoTo palorot2:
palorot1:        For b = 1 To 8 Step 1
        If x + 2 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next
        
          If Picture1(x + 3).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return


palorot2: If rot = 1 Then GoTo 50 Else GoTo pieza3
50:         For b = 1 To 8 Step 1
            If x = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
             Next
             If Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 31).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return
            
pieza3: If pieza = 3 Then GoTo 356 Else GoTo pieza4
356: If rot = 1 Then GoTo rot1pi3 Else GoTo rot2pi3
rot1pi3:   For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next
        
          If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return



rot2pi3: If rot = 2 Then GoTo 38 Else GoTo pieza4
38: For b = 1 To 8 Step 1
        If x + 17 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next
                  If Picture1(x + 18).BackColor = &H8000000F And Picture1(x + 1).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return



pieza4: If pieza = 4 Then GoTo pie4com Else GoTo pie5com
pie4com: If rot = 1 Then GoTo pie41com Else If rot = 2 Then GoTo pie42com
pie41com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return


pie42com: If rot = 2 Then GoTo rot22222 Else GoTo pie5com
rot22222: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return




pie5com: If pieza = 5 Then GoTo pie50com Else GoTo pi6com

pie50com: If rot = 1 Then GoTo pie51com Else GoTo pie52com
pie51com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return

pie52com: If rot = 2 Then GoTo pie522com Else GoTo pie53com

pie522com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return


pie53com: If rot = 3 Then GoTo pie533com Else GoTo pie54com

pie533com: For b = 1 To 8 Step 1
        If x = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return

pie54com: If rot = 4 Then GoTo pie544com Else GoTo pi6com
pie544com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 16).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return



pi6com: If pieza = 6 Then GoTo pi61com Else GoTo pi7com

pi61com: If rot = 1 Then GoTo pi611com Else GoTo pi62com
pi611com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 1).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 18).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return

pi62com: If rot = 2 Then GoTo pi622com Else GoTo pi63com
pi622com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 16).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return
pi63com: If rot = 3 Then GoTo pi633com Else GoTo pi64com
pi633com: For b = 1 To 8 Step 1
        If x = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return
pi64com: If rot = 4 Then GoTo pi644com Else GoTo pi7com
pi644com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 14).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return


pi7com: If pieza = 7 Then GoTo pi71com Else GoTo dibujar

pi71com: If rot = 1 Then GoTo pi711com Else GoTo pi72com
pi711com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return

pi72com: If rot = 2 Then GoTo pi722com Else GoTo pi73com
pi722com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return
pi73com: If rot = 3 Then GoTo pi733com Else GoTo pi74com
pi733com: For b = 1 To 8 Step 1
        If x = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 1).BackColor = &H8000000F And Picture1(x + 17).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return
pi74com: If rot = 4 Then GoTo pi744com Else GoTo dibujar
pi744com: For b = 1 To 8 Step 1
        If x + 1 = b * 16 Then caer = 1: GoSub dibujar: GoSub genera
        Next: If Picture1(x + 2).BackColor = &H8000000F And Picture1(x - 15).BackColor = &H8000000F Then GoSub borrar: x = x + 1 Else caer = 1: Return









dibujar:
If pieza = 1 Then Picture1(x).BackColor = &H80C0FF: Picture1(x + 1).BackColor = &H80C0FF: Picture1(x + 16).BackColor = &H80C0FF: Picture1(x + 17).BackColor = &H80C0FF: Return
If pieza = 2 And rot = 2 Then Picture1(x).BackColor = &HFF&: Picture1(x - 1).BackColor = &HFF&: Picture1(x + 1).BackColor = &HFF&: Picture1(x + 2).BackColor = &HFF&: Return
If pieza = 2 And rot = 1 Then Picture1(x - 32).BackColor = &HFF&: Picture1(x - 16).BackColor = &HFF&: Picture1(x).BackColor = &HFF&: Picture1(x + 16).BackColor = &HFF&: Return
If pieza = 3 And rot = 1 Then Picture1(x).BackColor = &HFF00&: Picture1(x + 16).BackColor = &HFF00&: Picture1(x + 1).BackColor = &HFF00&: Picture1(x - 15).BackColor = &HFF00&: Return
If pieza = 3 And rot = 2 Then Picture1(x).BackColor = &HFF00&: Picture1(x + 17).BackColor = &HFF00&: Picture1(x + 16).BackColor = &HFF00&: Picture1(x - 1).BackColor = &HFF00&: Return
If pieza = 4 And rot = 1 Then Picture1(x).BackColor = &HFF00FF: Picture1(x + 17).BackColor = &HFF00FF: Picture1(x + 1).BackColor = &HFF00FF: Picture1(x - 16).BackColor = &HFF00FF: Return
If pieza = 4 And rot = 2 Then Picture1(x).BackColor = &HFF00FF: Picture1(x - 1).BackColor = &HFF00FF: Picture1(x - 16).BackColor = &HFF00FF: Picture1(x - 15).BackColor = &HFF00FF: Return
If pieza = 5 And rot = 1 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 16).BackColor = &HFF0000: Picture1(x - 16).BackColor = &HFF0000: Picture1(x - 15).BackColor = &HFF0000: Return
If pieza = 5 And rot = 2 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 17).BackColor = &HFF0000: Picture1(x - 1).BackColor = &HFF0000: Picture1(x + 1).BackColor = &HFF0000: Return
If pieza = 5 And rot = 3 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 16).BackColor = &HFF0000: Picture1(x - 16).BackColor = &HFF0000: Picture1(x + 15).BackColor = &HFF0000: Return
If pieza = 5 And rot = 4 Then Picture1(x).BackColor = &HFF0000: Picture1(x + 1).BackColor = &HFF0000: Picture1(x - 1).BackColor = &HFF0000: Picture1(x - 17).BackColor = &HFF0000: Return
If pieza = 6 And rot = 1 Then Picture1(x).BackColor = &H404080: Picture1(x + 16).BackColor = &H404080: Picture1(x - 16).BackColor = &H404080: Picture1(x + 17).BackColor = &H404080: Return
If pieza = 6 And rot = 4 Then Picture1(x).BackColor = &H404080: Picture1(x + 1).BackColor = &H404080: Picture1(x - 1).BackColor = &H404080: Picture1(x - 15).BackColor = &H404080: Return
If pieza = 6 And rot = 3 Then Picture1(x).BackColor = &H404080: Picture1(x + 16).BackColor = &H404080: Picture1(x - 16).BackColor = &H404080: Picture1(x - 17).BackColor = &H404080: Return
If pieza = 6 And rot = 2 Then Picture1(x).BackColor = &H404080: Picture1(x + 1).BackColor = &H404080: Picture1(x - 1).BackColor = &H404080: Picture1(x + 15).BackColor = &H404080: Return
If pieza = 7 And rot = 1 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0: Return
If pieza = 7 And rot = 2 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0: Picture1(x - 1).BackColor = &H3DFAF0: Return
If pieza = 7 And rot = 3 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x - 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x + 16).BackColor = &H3DFAF0: Return
If pieza = 7 And rot = 4 Then Picture1(x).BackColor = &H3DFAF0: Picture1(x + 1).BackColor = &H3DFAF0: Picture1(x - 16).BackColor = &H3DFAF0: Picture1(x - 1).BackColor = &H3DFAF0: Return






genera: lineaseg = 0: completa = 1
     For i = 15 To 1 Step -1
    For a = 1 To 113 Step 16: completa = 1
   

        If Picture1(a + i).BackColor = &H8000000F Then completa = 0: GoTo 6060
    Next a
6060: If completa = 1 Then GoTo 6061 Else GoTo 6070
6061 lineaseg = lineaseg + 1: lineas = lineas + 1:
If lineaseg = 1 Then z = 1: p1 = i
If lineaseg = 2 Then z = 2: p2 = i
If lineaseg = 3 Then z = 3: p3 = i
If lineaseg = 4 Then z = 4


6070 Next
If lineaseg = 1 Then puntuacion = puntuacion + 100: GoTo li1
If lineaseg = 2 Then puntuacion = puntuacion + 200: GoTo li2
If lineaseg = 3 Then puntuacion = puntuacion + 300: GoTo li3
If lineaseg = 4 Then puntuacion = puntuacion + 500: GoTo li4

li1: For a = 1 To 113 Step 16
 For g = p1 To 1 Step -1

Picture1(a + g).BackColor = Picture1(a - z + g).BackColor
Next g
Next a
Picture1(1).BackColor = &H8000000F:  Picture1(17).BackColor = &H8000000F: Picture1(33).BackColor = &H8000000F: Picture1(49).BackColor = &H8000000F: Picture1(65).BackColor = &H8000000F: Picture1(81).BackColor = &H8000000F: Picture1(97).BackColor = &H8000000F: Picture1(113).BackColor = &H8000000F
GoTo 33
li2: For a = 1 To 113 Step 16
 For g = p1 To 2 Step -1

Picture1(a + g).BackColor = Picture1(a - 1 + g).BackColor
Next g
Next a

For b = 1 To 113 Step 16
 For c = p2 + 1 To 2 Step -1

Picture1(b + c).BackColor = Picture1(b - 1 + c).BackColor
Next c
Next b
Picture1(1).BackColor = &H8000000F:  Picture1(17).BackColor = &H8000000F: Picture1(33).BackColor = &H8000000F: Picture1(49).BackColor = &H8000000F: Picture1(65).BackColor = &H8000000F: Picture1(81).BackColor = &H8000000F: Picture1(97).BackColor = &H8000000F: Picture1(113).BackColor = &H8000000F
GoTo 33




li3: For a = 1 To 113 Step 16
 For g = p1 To 2 Step -1

Picture1(a + g).BackColor = Picture1(a - 1 + g).BackColor
Next g
Next a

For b = 1 To 113 Step 16
 For c = p2 + 1 To 2 Step -1

Picture1(b + c).BackColor = Picture1(b - 1 + c).BackColor
Next c
Next b

For b = 1 To 113 Step 16
 For c = p3 + 2 To 2 Step -1

Picture1(b + c).BackColor = Picture1(b - 1 + c).BackColor
Next c
Next b


Picture1(1).BackColor = &H8000000F:  Picture1(17).BackColor = &H8000000F: Picture1(33).BackColor = &H8000000F: Picture1(49).BackColor = &H8000000F: Picture1(65).BackColor = &H8000000F: Picture1(81).BackColor = &H8000000F: Picture1(97).BackColor = &H8000000F: Picture1(113).BackColor = &H8000000F
GoTo 33
li4: For a = 1 To 113 Step 16
 For g = p1 To 4 Step -1

Picture1(a + g).BackColor = Picture1(a - z + g).BackColor
Next g
Next a
Picture1(1).BackColor = &H8000000F:  Picture1(17).BackColor = &H8000000F: Picture1(33).BackColor = &H8000000F: Picture1(49).BackColor = &H8000000F: Picture1(65).BackColor = &H8000000F: Picture1(81).BackColor = &H8000000F: Picture1(97).BackColor = &H8000000F: Picture1(113).BackColor = &H8000000F
GoTo 33

33: Lpuntuacion = puntuacion
If lineas > 9 And lineas < 20 Then nivel = 1
If lineas > 19 And lineas < 30 Then nivel = 2
If lineas > 29 And lineas < 40 Then nivel = 3
If lineas > 39 And lineas < 50 Then nivel = 4
If lineas > 49 And lineas < 60 Then nivel = 5
If lineas > 59 And lineas < 70 Then nivel = 6
If lineas > 69 And lineas < 80 Then nivel = 7
If lineas > 79 And lineas < 90 Then nivel = 8
If lineas > 89 And lineas < 100 Then nivel = 9
If lineas > 99 And lineas < 110 Then nivel = 10
If lineas > 109 And lineas < 120 Then nivel = 11
If lineas > 119 And lineas < 130 Then nivel = 12


Lnivel = nivel
Llineas = lineas
lrecord = record








x = 49:  pieza = piezasig: piezasig = Int(Rnd(1) * 7) + 1:  caer = 0: rot = 1: If pieza = 10 Then x = 50:: Return
If piezasig = 1 Then Picture2.Picture = Pieza1.Picture
If piezasig = 2 Then Picture2.Picture = Pieza2.Picture
If piezasig = 3 Then Picture2.Picture = pieza3.Picture
If piezasig = 4 Then Picture2.Picture = pieza4.Picture
If piezasig = 5 Then Picture2.Picture = pieza5.Picture
If piezasig = 6 Then Picture2.Picture = pieza6.Picture
If piezasig = 7 Then Picture2.Picture = pieza7.Picture





fin:

End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
nueva.Enabled = True


Form2.Show (1)



End Sub
