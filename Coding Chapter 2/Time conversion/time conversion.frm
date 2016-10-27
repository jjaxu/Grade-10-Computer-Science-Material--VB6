VERSION 5.00
Begin VB.Form timeconversion 
   Caption         =   "Time Conversion"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRight 
      Caption         =   "->"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<-"
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
      Left            =   2520
      Picture         =   "time conversion.frx":0000
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtDays 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtTM 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtMinutes 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtHours 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"time conversion.frx":AF9D62
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Label lblDays 
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTM 
      Caption         =   "Total Minutes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblHours 
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "timeconversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: Ocbober 2 2013
'Purpose: Play around with multiple dims and commands
Option Explicit

Private Sub cmdLeft_Click()
Dim IntDays As Integer
Dim IntHours As Integer
Dim IntMinutes As Integer
Dim IntTM As Integer
'initialize
IntDays = 0
IntHours = 0
IntMinutes = 0
IntTM = 0
'input
IntDays = Val(txtDays.Text)
IntHours = Val(txtHours.Text)
IntMinutes = Val(txtMinutes.Text)
IntTM = Val(txtTM.Text)
'process/cal
IntDays = IntTM \ 1440
IntHours = (IntTM - IntDays * 1440) \ 60
IntMinutes = (IntTM - (IntDays * 1440 + IntHours * 60)) Mod 60
'output
txtDays.Text = IntDays
txtHours.Text = IntHours
txtMinutes.Text = IntMinutes
End Sub

Private Sub cmdRight_Click()
'declare
Dim IntDays As Single
Dim IntHours As Single
Dim IntMinutes As Single
Dim IntTM As Single
'initialize
IntDays = 0
IntHours = 0
IntMinutes = 0
IntTM = 0
'input
IntDays = Val(txtDays.Text)
IntHours = Val(txtHours.Text)
IntMinutes = Val(txtMinutes.Text)
IntTM = Val(txtTM.Text)
'process/cal
IntTM = IntDays * 1440 + IntHours * 60 + IntMinutes
'output
txtTM.Text = IntTM
End Sub

