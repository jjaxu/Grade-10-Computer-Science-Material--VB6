VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStop 
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtStart 
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   8055
   End
   Begin VB.ListBox lstOutput 
      Height          =   3180
      Left            =   5880
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblStop 
      Caption         =   "Stop"
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'
Option Explicit

Private Sub cmdCount_Click()
'Declare
Dim intStart As Integer
Dim intStop As Integer
Dim intCount As Integer
'Initialize
intCount = 0
'Input
intStart = Val(txtStart.Text)
intStop = Val(txtStop.Text)
'Process / Output
intCount = intStart
lstOutput.Clear
Do While (intCount <= intStop)
    lstOutput.AddItem intCount
    intCount = intCount + 1
Loop

End Sub
