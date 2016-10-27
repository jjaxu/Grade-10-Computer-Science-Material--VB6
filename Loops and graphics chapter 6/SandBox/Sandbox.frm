VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   6120
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   480
      ScaleHeight     =   365
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   581
      TabIndex        =   0
      Top             =   360
      Width           =   8775
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
'GV
Dim intX1 As Integer
Dim intY1 As Integer
Dim intX2 As Integer
Dim intY2 As Integer

Dim intCounter As Integer

Private Sub Command1_Click()
Picture1.Print "(" & intX1 & ", " & intY1 & ")"
Picture1.Print "(" & intX2 & ", " & intY2 & ")"
End Sub

Private Sub Form_Load()
intX1 = 0
intY2 = 0
intX1 = 0
intY2 = 0
intCounter = 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Declare
'Initialize
'Process
intCounter = intCounter + 1
'


'Picture1.Print "(" & X & ", " & Y & ")"
If intCounter <= 2 Then
    Picture1.PSet (X, Y), vbRed
    Picture1.Circle (X, Y), 5, vbBlue
    If intCounter = 1 Then
        intX1 = X
        intY1 = Y
    Else
        intX2 = X
        intY2 = Y
    End If
Else
    MsgBox ("Stop!!!")
    
End If

'Picture1.Print Button
End Sub
