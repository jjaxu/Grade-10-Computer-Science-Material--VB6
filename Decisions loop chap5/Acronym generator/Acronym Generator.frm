VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   6615
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   6615
   End
   Begin VB.Label lblOutput 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   6615
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      Caption         =   "You Acronym:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   6495
   End
   Begin VB.Label lblTitle 
      Caption         =   "Acronym Generator"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: November 30 2013
'Purpose: More practice with loops with strings
Option Explicit

Private Sub cmdGenerate_Click()
'Declare
Dim strInput As String
Dim strResult As String
Dim strNewPhrase As String
Dim intCounter As Integer
'Initialize
strInput = ""
strResult = ""
strNewPhrase = ""
intCounter = 0
'Input
strInput = Trim(UCase(txtInput.Text))
'Process / Calculation / output
strResult = Left(strInput, 1)

Do While InStr(strInput, " ") > 0
    intCounter = InStr(strInput, " ")
    strNewPhrase = Mid(strInput, intCounter + 1)
    strResult = Trim(strResult & Left(strNewPhrase, 1))
    strInput = strNewPhrase
Loop
lblOutput.Caption = strResult

End Sub
