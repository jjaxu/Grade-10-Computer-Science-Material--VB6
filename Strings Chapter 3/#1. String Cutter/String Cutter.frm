VERSION 5.00
Begin VB.Form StringCutter 
   Caption         =   "String Cutter"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      ScaleHeight     =   3915
      ScaleWidth      =   7155
      TabIndex        =   3
      Top             =   4320
      Width           =   7215
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cut me up!!!"
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
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   7215
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
      Height          =   915
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter at least 10 characters."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblMessage 
      Caption         =   "String Cutter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "StringCutter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: October 16 2013
'Purpose: Play around with strings
Option Explicit

Private Sub Form_Load()
cmdCut.Enabled = False
End Sub

Private Sub cmdCut_Click()
'Declare
Dim strInput As String
Dim strA As String
Dim strB As String
Dim strC As String
Dim strD As String
Dim strE As String
Dim strF As String
Dim strG As String
Dim strH As String
Dim strI As String
Dim strJ As String

'Initialize
strInput = ""
strA = ""
strB = ""
strB = ""
strD = ""
strE = ""
strF = ""
strG = ""
strH = ""
strI = ""
strJ = ""

'Input
strInput = txtInput.Text

'Calculation/process
strA = Left(strInput, 1)
strB = Right(strInput, 1)
strC = Left(strInput, 3)
strD = Right(strInput, 2)
strE = Mid(strInput, 6)
strF = Mid(strInput, 3, 4)
strG = Mid(strInput, 4, 2)
strH = Mid(strInput, 5, 1)
strI = Mid(strInput, 3, 1)
strJ = Left(strInput, Len(strInput) - 3)

'Output
picOutput.Cls
picOutput.Print "The first character is: "; strA
picOutput.Print "The last character is: "; strB
picOutput.Print "The first 2 characters are: "; strC
picOutput.Print "The last 2 characters are: "; strD
picOutput.Print "All but the first 5 Characters are: "; strE
picOutput.Print "The 4 characters starting in the 3rd are: "; strF
picOutput.Print "The 2 characters starting in the 4th are: "; strG
picOutput.Print "The 5th characters is: "; strH
picOutput.Print "The 3rd characters is: "; strI
picOutput.Print "All but the last 3 characters are: "; strJ
End Sub

Private Sub txtInput_Change()
If txtInput.Text = "" Then cmdCut.Enabled = False
If txtInput.Text <> "" Then cmdCut.Enabled = True
If Len(txtInput.Text) < 10 Then cmdCut.Enabled = False
picOutput.Cls
End Sub
