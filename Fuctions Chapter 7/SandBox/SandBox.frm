VERSION 5.00
Begin VB.Form frmSubP 
   Caption         =   "Sub Programming"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumb2 
      Height          =   975
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtNumb1 
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdDemo 
      Caption         =   "Demo"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   5160
      ScaleHeight     =   4875
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label lblOutput 
      Caption         =   "Label1"
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   4575
   End
End
Attribute VB_Name = "frmSubP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'
Option Explicit
'GV

Private Sub cmdDemo_Click()
'Declre
Dim intNumb1 As Integer
Dim intNumb2 As Integer
Dim intArea As Integer
'Init
'Input
intNumb1 = Val(txtNumb1.Text)
intNumb2 = Val(txtNumb2.Text)
'Process
intArea = calArea(intNumb1, intNumb2)
'Output
picOutput.Cls
picOutput.Print "The Area is : " & intArea
End Sub

Public Function calArea(intFirst As Integer, intSecond As Integer)
'Dec
Dim intAnswer As Integer
'Init
'Input
'Process
intAnswer = intFirst * intSecond
'Output
calArea = intAnswer
End Function

