VERSION 5.00
Begin VB.Form frmWordSorter 
   Caption         =   "Word Sorter"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   7635
      TabIndex        =   8
      Top             =   5520
      Width           =   7695
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   7695
   End
   Begin VB.TextBox txtWord3 
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
      Left            =   3720
      TabIndex        =   6
      Top             =   3480
      Width           =   4095
   End
   Begin VB.TextBox txtWord2 
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
      Left            =   3720
      TabIndex        =   5
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox txtWord1 
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
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblWord3 
      Caption         =   "Enter 3rd word:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblWord2 
      Caption         =   "Enter 2nd word:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label lblWord1 
      Caption         =   "Enter 1st word:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      Caption         =   "Word Sorter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   42
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
      Width           =   7455
   End
End
Attribute VB_Name = "frmWordSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: November 6 2013
'Purpose: Introduction to decisions
Option Explicit

Private Sub cmdSort_Click()
'declare
Dim strWord1 As String
Dim strWord2 As String
Dim strWord3 As String

'initialize
strWord1 = ""
strWord2 = ""
strWord3 = ""

'input
strWord1 = LCase(Trim(txtWord1.Text))
strWord2 = LCase(Trim(txtWord2.Text))
strWord3 = LCase(Trim(txtWord3.Text))

'process / output
picOutput.Cls
If strWord1 <= strWord2 And strWord2 <= strWord3 And strWord1 <= strWord3 Then
    picOutput.Print strWord1 & "," & strWord2 & "," & strWord3
ElseIf strWord1 <= strWord2 And strWord2 >= strWord3 And strWord1 <= strWord3 Then
    picOutput.Print strWord1 & "," & strWord3 & "," & strWord2

ElseIf strWord2 <= strWord1 And strWord1 <= strWord3 And strWord2 <= strWord3 Then
    picOutput.Print strWord2 & "," & strWord1 & "," & strWord3
ElseIf strWord2 <= strWord1 And strWord1 >= strWord3 And strWord2 <= strWord3 Then
    picOutput.Print strWord2 & "," & strWord3 & "," & strWord1

ElseIf strWord3 <= strWord1 And strWord1 <= strWord2 And strWord3 <= strWord2 Then
    picOutput.Print strWord3 & "," & strWord1 & "," & strWord2
ElseIf strWord3 <= strWord1 And strWord1 >= strWord2 And strWord3 <= strWord2 Then
    picOutput.Print strWord3 & "," & strWord2 & "," & strWord1

End If
End Sub
