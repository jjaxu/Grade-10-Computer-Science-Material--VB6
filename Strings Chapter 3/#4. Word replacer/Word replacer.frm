VERSION 5.00
Begin VB.Form WordReplacer 
   Caption         =   "Word Replacer"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace word"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   240
      TabIndex        =   8
      Top             =   5280
      Width           =   7575
   End
   Begin VB.TextBox txtReplace 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   6
      Top             =   4200
      Width           =   4095
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   7695
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace with:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label lblFind 
      Caption         =   "Find What:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblSentance 
      Caption         =   "Enter your sentance here:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblMessage 
      Caption         =   $"Word replacer.frx":0000
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Word Replacer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "WordReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie
'Date: October 20 2013
'Purpose: Play around with more strings and search and replace strings
Option Explicit

Private Sub cmdReplace_Click()
'Declare
Dim strInput As String
Dim strFind As String
Dim strReplace As String
Dim strOutput As String
Dim strBefore As String
Dim strAfter As String
 Dim intCharacters As Integer

'Initialize
strInput = ""
strFind = ""
strReplace = ""
strOutput = ""
strBefore = ""
strAfter = ""
intCharacters = 0

'Input
strInput = txtInput.Text
strFind = txtFind.Text
strReplace = txtReplace.Text
strOutput = txtInput.Text

'Process/calculation
intCharacters = InStr(strInput, strFind)

strBefore = Left(strInput, intCharacters - 1)
strAfter = Mid(strInput, Len(strFind & strBefore) + 1)

'Output
txtInput.Text = strBefore & strReplace & strAfter
End Sub

