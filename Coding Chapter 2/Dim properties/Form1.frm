VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtwidth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   6495
   End
   Begin VB.TextBox txtlength 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   6495
   End
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "Calculate Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   10935
   End
   Begin VB.Label lbloutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   10815
   End
   Begin VB.Label lblwidth 
      Caption         =   "Enter Width:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lbllength 
      Caption         =   "Enter Length:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu'
'Date:'
'purpose'
Option Explicit

Private Sub cmdcalculate_Click()
'delcare
Dim SglLength As Single
Dim SglWidth As Single
Dim SglArea As Single
'Initialize
SglArea = 0
SglLength = 0
SglWidth = 0
'input
SglLength = Val(txtlength.Text)
SglWidth = Val(txtwidth.Text)
'process/calculation
SglArea = SglLength * SglWidth
'output
lbloutput.Caption = "The area of this box is " & SglArea & " Sq. Units"
End Sub

Private Sub Form_Load()

End Sub
