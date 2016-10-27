VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text Changer"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Style 
      Caption         =   "Style"
      Height          =   5775
      Left            =   9840
      TabIndex        =   9
      Top             =   1320
      Width           =   2295
      Begin VB.OptionButton styleitalic 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   12
         Top             =   4080
         Width           =   1455
      End
      Begin VB.OptionButton stylebold 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
      End
      Begin VB.OptionButton stylenormal 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame caption 
      Caption         =   "Caption"
      Height          =   5775
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   6135
      Begin VB.OptionButton text3 
         Caption         =   "Have a nice day!"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Width           =   5415
      End
      Begin VB.OptionButton text2 
         Caption         =   "How are you?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   5415
      End
      Begin VB.OptionButton text1 
         Caption         =   "Hello, my name is Jackie."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   5535
      End
   End
   Begin VB.Frame font 
      Caption         =   "Font"
      Height          =   5775
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
      Begin VB.OptionButton fontsimplex 
         Caption         =   "Simplex"
         BeginProperty Font 
            Name            =   "Simplex"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   5
         Top             =   4320
         Width           =   2535
      End
      Begin VB.OptionButton fonttnr 
         Caption         =   "Terminal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   2295
      End
      Begin VB.OptionButton fontcalibri 
         Caption         =   "Calibri"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Label lblmessage 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Choose a Caption, Font, and Style, have fun!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11865
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fontsimplex_Click()
lblmessage.font = "simplex"
If lblmessage = "Choose a Caption, Font, and Style, have fun!" Then lblmessage = ""
End Sub

Private Sub fontcalibri_Click()
lblmessage.font = "calibri"
End Sub

Private Sub fonttnr_Click()
lblmessage.font = "terminal"
If lblmessage = "Choose a Caption, Font, and Style, have fun!" Then lblmessage = ""
End Sub


Private Sub Form_Load()
'Author: Jackie Xu'
'Date: September 20 2013'
End Sub

Private Sub stylebold_Click()
lblmessage.FontBold = True
lblmessage.FontItalic = False
If lblmessage = "Choose a Caption, Font, and Style, have fun!" Then lblmessage = ""
End Sub

Private Sub styleitalic_Click()
lblmessage.FontItalic = True
lblmessage.FontBold = False
If lblmessage = "Choose a Caption, Font, and Style, have fun!" Then lblmessage = ""
End Sub

Private Sub stylenormal_Click()
lblmessage.FontBold = False
lblmessage.FontItalic = False
If lblmessage = "Choose a Caption, Font, and Style, have fun!" Then lblmessage = ""
End Sub

Private Sub text1_Click()
lblmessage.caption = "Hello, my name is Jackie."
End Sub

Private Sub text2_Click()
lblmessage.caption = "How are you?"
End Sub


Private Sub text3_Click()
lblmessage.caption = "Have a nice day!"
End Sub
