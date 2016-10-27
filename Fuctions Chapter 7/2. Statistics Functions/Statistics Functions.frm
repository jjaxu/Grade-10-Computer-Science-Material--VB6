VERSION 5.00
Begin VB.Form frmStaticstic 
   Caption         =   "Statistics Functions"
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   5
      Top             =   1440
      Width           =   6255
   End
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   7800
      Width           =   6255
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   6255
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Current Status"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   6255
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Label lblTitle 
      Caption         =   "Statistics Functions"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
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
      Width           =   6015
   End
End
Attribute VB_Name = "frmStaticstic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: December 18 2013
'Purpose: More functions with subs and objects
Option Explicit
'GV
Dim sglInput As Single

Private Sub cmdAdd_Click()

'Initialize
sglInput = 0
'Input
sglInput = Trim(Val(txtInput.Text))
'Process
picOutput.Cls
If sglInput <= 100 And sglInput > 0 Then
lstOutput.AddItem sglInput
End If
'Output
End Sub

Private Sub cmdStatus_Click()
picOutput.Cls
picOutput.Print "Smallest Number: " & findSmallest
picOutput.Print "Largest Number: " & findLargest
picOutput.Print "Total Value: " & findTotal
picOutput.Print "Average Value: " & findAverage
End Sub

'Functions
Public Function findSmallest()
'Declare
Dim sglSmallest As Single
Dim intlistcount As Integer
'Process
sglSmallest = Val(lstOutput.List(0))
For intlistcount = 0 To lstOutput.ListCount - 1
    If lstOutput.List(intlistcount) < sglSmallest Then
        sglSmallest = lstOutput.List(intlistcount)
    End If
Next intlistcount
findSmallest = sglSmallest
End Function

Public Function findLargest()
'Dec
Dim sglLargest As Single
Dim intlistcount As Integer
'Init
sglLargest = 0
'Process
For intlistcount = 0 To lstOutput.ListCount - 1
    If lstOutput.List(intlistcount) > sglLargest Then
        sglLargest = lstOutput.List(intlistcount)
    End If
Next intlistcount
findLargest = sglLargest
End Function

Public Function findTotal()
'Dec
Dim sglTotal As Single
Dim intlistcount As Integer
'Process
For intlistcount = 0 To lstOutput.ListCount - 1
    sglTotal = sglTotal + lstOutput.List(intlistcount)
Next intlistcount
findTotal = sglTotal
End Function

Public Function findAverage()
'Dim
Dim sglAverage As Single
'Process
sglAverage = findTotal / lstOutput.ListCount
findAverage = sglAverage
End Function

