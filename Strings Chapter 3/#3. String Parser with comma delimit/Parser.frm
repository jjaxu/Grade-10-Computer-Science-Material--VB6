VERSION 5.00
Begin VB.Form StringParser 
   Caption         =   "Sting Parser"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdParse 
      Caption         =   "Click me to Parse the string"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   7815
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label lblString3 
      Caption         =   "String 3:"
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
      TabIndex        =   9
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label lblString2 
      Caption         =   "String 2:"
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
      TabIndex        =   8
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblString1 
      Caption         =   "String 1:"
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
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblOutput3 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   6
      Top             =   5880
      Width           =   4815
   End
   Begin VB.Label lblOutput2 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   5
      Top             =   4800
      Width           =   4815
   End
   Begin VB.Label lblOutput1 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   4
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Label lblMessage 
      Caption         =   "Enter 3 pieces of strings and separate them by 2 commas, then click ""Parse"" to split into 3 parts."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      Caption         =   "String Parser"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "StringParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: December 19 2013
'Purpose: Play around with more strings and as well as spliting strings
Option Explicit

Private Sub CmdParse_Click()
'Declare
Dim strInput As String
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim intComma1 As Integer
Dim strPhrase2 As String
Dim intComma2 As Integer

'Initialize
strInput = ""
str1 = ""
str2 = ""
str3 = ""
intComma1 = 0
strPhrase2 = ""
intComma2 = 0

'Input
strInput = Trim(txtInput.Text)

'Process/calculation
str1 = Left(strInput, InStr(strInput, ",") - 1)
intComma1 = InStr(strInput, ",")
strPhrase2 = Mid(strInput, intComma1 + 1)
intComma2 = InStr(strPhrase2, ",")
str2 = Left(strPhrase2, intComma2 - 1)
str3 = Mid(strPhrase2, intComma2 + 1)

'Output
lblOutput1 = str1
lblOutput2 = str2
lblOutput3 = str3
End Sub

Private Sub Form_Load()
CmdParse.Enabled = False
End Sub

Private Sub txtInput_Change()
Dim strInput As String
Dim strPhrase2 As String
strInput = ""
strPhrase2 = ""
strInput = txtInput.Text
If txtInput.Text = "" Then CmdParse.Enabled = False
If txtInput.Text <> "" Then CmdParse.Enabled = True
End Sub
