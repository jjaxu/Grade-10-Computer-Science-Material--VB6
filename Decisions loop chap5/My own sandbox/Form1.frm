VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   5880
      TabIndex        =   3
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   3135
      Left            =   960
      TabIndex        =   2
      Top             =   4200
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sglC As Single
Dim sglF As Single
Dim sglS As Single
Dim sglE As Single
Private Sub Command1_Click()


sglC = 0
sglF = 0
sglS = 0
sglE = 0

sglS = Val(Text1.Text)
sglE = Val(Text2.Text)

List1.Clear
If sglS > sglE Then
    List1.AddItem "Invalid Values"
Else
Do
sglF = 9 / 5 * sglS + 32
List1.AddItem sglS & "C" & " = " & sglF & "F"
sglS = sglS + 0.5
Loop While sglS <= sglE
End If

End Sub

