VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sorting SandBox"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   7215
   End
   Begin VB.TextBox txtSecond 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtFirst 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblOutput 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   7335
   End
   Begin VB.Label lblSecond 
      Caption         =   "Second number"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblFirst 
      Caption         =   "First number"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author
'Date
'Purpose
Option Explicit

Private Sub cmdSort_Click()
'Declare
Dim intFirst As Integer
Dim intSecond As Integer
Dim strAnswer As String
'Initialize
strAnswer = ""
'Input
intFirst = Val(txtFirst.Text)
intSecond = Val(txtSecond.Text)
'Process
If intFirst > intSecond Then
    strAnswer = intFirst & " is bigger than " & intSecond
ElseIf intFirst < intSecond Then
    strAnswer = intFirst & " is smaller than " & intSecond
Else
    strAnswer = intFirst & " is equal to " & intSecond
End If
'Output
lblOutput.Caption = strAnswer

End Sub
