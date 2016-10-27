VERSION 5.00
Begin VB.Form frmDominoes 
   Caption         =   "Dominoes"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   702
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDom2R 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   2880
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   7
      Top             =   4920
      Width           =   2655
   End
   Begin VB.PictureBox picDom2L 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   240
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   6
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   9360
      Width           =   5415
   End
   Begin VB.PictureBox picDom1R 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   2880
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.PictureBox picDom1L 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   240
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblJoin 
      Caption         =   "Joinable:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   8520
      Width           =   5295
   End
   Begin VB.Label lblDom2R 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label lblDom2L 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label lblDom1R 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblDom1L 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Caption         =   "Dominoes"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmDominoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: December 12 2013
'Purpose: To learn how to use Functions
Option Explicit
'GV
Dim intDomL1 As Integer
Dim intDomR1 As Integer
Dim intDomL2 As Integer
Dim intDomR2 As Integer

Private Sub cmdGenerate_Click()
'Initialize
intDomL1 = 0
intDomR1 = 0

intDomL2 = 0
intDomR2 = 0
'input / No input
'Process
intDomL1 = leftSide
intDomR1 = rightSide

intDomL2 = leftSide
intDomR2 = rightSide

picDom1L.Cls
picDom2L.Cls
picDom1R.Cls
picDom2R.Cls
'DOM LEFT 1
If intDomL1 = 1 Then
    picDom1L.Circle (88.5, 84.5), 10, vbBlack
ElseIf intDomL1 = 2 Then
    picDom1L.Circle (14, 14), 10, vbBlack
    picDom1L.Circle (154, 150), 10, vbBlack
ElseIf intDomL1 = 3 Then
    picDom1L.Circle (88.5, 84.5), 10, vbBlack
    picDom1L.Circle (14, 14), 10, vbBlack
    picDom1L.Circle (154, 150), 10, vbBlack
ElseIf intDomL1 = 4 Then
    picDom1L.Circle (14, 14), 10, vbBlack
    picDom1L.Circle (14, 150), 10, vbBlack
    picDom1L.Circle (154, 14), 10, vbBlack
    picDom1L.Circle (154, 150), 10, vbBlack
ElseIf intDomL1 = 5 Then
    picDom1L.Circle (88.5, 84.5), 10, vbBlack
    picDom1L.Circle (14, 14), 10, vbBlack
    picDom1L.Circle (14, 150), 10, vbBlack
    picDom1L.Circle (154, 14), 10, vbBlack
    picDom1L.Circle (154, 150), 10, vbBlack
ElseIf intDomL1 = 6 Then
    picDom1L.Circle (14, 14), 10, vbBlack
    picDom1L.Circle (88.5, 14), 10, vbBlack
    picDom1L.Circle (14, 150), 10, vbBlack
    picDom1L.Circle (154, 14), 10, vbBlack
    picDom1L.Circle (88.5, 150), 10, vbBlack
    picDom1L.Circle (154, 150), 10, vbBlack
End If

'DOM RIGHT 1
If intDomR1 = 1 Then
    picDom1R.Circle (88.5, 84.5), 10, vbBlack
ElseIf intDomR1 = 2 Then
    picDom1R.Circle (14, 14), 10, vbBlack
    picDom1R.Circle (154, 150), 10, vbBlack
ElseIf intDomR1 = 3 Then
    picDom1R.Circle (88.5, 84.5), 10, vbBlack
    picDom1R.Circle (14, 14), 10, vbBlack
    picDom1R.Circle (154, 150), 10, vbBlack
ElseIf intDomR1 = 4 Then
    picDom1R.Circle (14, 14), 10, vbBlack
    picDom1R.Circle (14, 150), 10, vbBlack
    picDom1R.Circle (154, 14), 10, vbBlack
    picDom1R.Circle (154, 150), 10, vbBlack
ElseIf intDomR1 = 5 Then
    picDom1R.Circle (88.5, 84.5), 10, vbBlack
    picDom1R.Circle (14, 14), 10, vbBlack
    picDom1R.Circle (14, 150), 10, vbBlack
    picDom1R.Circle (154, 14), 10, vbBlack
    picDom1R.Circle (154, 150), 10, vbBlack
ElseIf intDomR1 = 6 Then
    picDom1R.Circle (14, 14), 10, vbBlack
    picDom1R.Circle (88.5, 14), 10, vbBlack
    picDom1R.Circle (14, 150), 10, vbBlack
    picDom1R.Circle (154, 14), 10, vbBlack
    picDom1R.Circle (88.5, 150), 10, vbBlack
    picDom1R.Circle (154, 150), 10, vbBlack
End If

'DOM LEFT 2
If intDomL2 = 1 Then
    picDom2L.Circle (88.5, 84.5), 10, vbBlack
ElseIf intDomL2 = 2 Then
    picDom2L.Circle (14, 14), 10, vbBlack
    picDom2L.Circle (154, 150), 10, vbBlack
ElseIf intDomL2 = 3 Then
    picDom2L.Circle (88.5, 84.5), 10, vbBlack
    picDom2L.Circle (14, 14), 10, vbBlack
    picDom2L.Circle (154, 150), 10, vbBlack
ElseIf intDomL2 = 4 Then
    picDom2L.Circle (14, 14), 10, vbBlack
    picDom2L.Circle (14, 150), 10, vbBlack
    picDom2L.Circle (154, 14), 10, vbBlack
    picDom2L.Circle (154, 150), 10, vbBlack
ElseIf intDomL2 = 5 Then
    picDom2L.Circle (88.5, 84.5), 10, vbBlack
    picDom2L.Circle (14, 14), 10, vbBlack
    picDom2L.Circle (14, 150), 10, vbBlack
    picDom2L.Circle (154, 14), 10, vbBlack
    picDom2L.Circle (154, 150), 10, vbBlack
ElseIf intDomL2 = 6 Then
    picDom2L.Circle (14, 14), 10, vbBlack
    picDom2L.Circle (88.5, 14), 10, vbBlack
    picDom2L.Circle (14, 150), 10, vbBlack
    picDom2L.Circle (154, 14), 10, vbBlack
    picDom2L.Circle (88.5, 150), 10, vbBlack
    picDom2L.Circle (154, 150), 10, vbBlack
End If

'DOM RIGHT 2
If intDomR2 = 1 Then
    picDom2R.Circle (88.5, 84.5), 10, vbBlack
ElseIf intDomR2 = 2 Then
    picDom2R.Circle (14, 14), 10, vbBlack
    picDom2R.Circle (154, 150), 10, vbBlack
ElseIf intDomR2 = 3 Then
    picDom2R.Circle (88.5, 84.5), 10, vbBlack
    picDom2R.Circle (14, 14), 10, vbBlack
    picDom2R.Circle (154, 150), 10, vbBlack
ElseIf intDomR2 = 4 Then
    picDom2R.Circle (14, 14), 10, vbBlack
    picDom2R.Circle (14, 150), 10, vbBlack
    picDom2R.Circle (154, 14), 10, vbBlack
    picDom2R.Circle (154, 150), 10, vbBlack
ElseIf intDomR2 = 5 Then
    picDom2R.Circle (88.5, 84.5), 10, vbBlack
    picDom2R.Circle (14, 14), 10, vbBlack
    picDom2R.Circle (14, 150), 10, vbBlack
    picDom2R.Circle (154, 14), 10, vbBlack
    picDom2R.Circle (154, 150), 10, vbBlack
ElseIf intDomR2 = 6 Then
    picDom2R.Circle (14, 14), 10, vbBlack
    picDom2R.Circle (88.5, 14), 10, vbBlack
    picDom2R.Circle (14, 150), 10, vbBlack
    picDom2R.Circle (154, 14), 10, vbBlack
    picDom2R.Circle (88.5, 150), 10, vbBlack
    picDom2R.Circle (154, 150), 10, vbBlack
End If


'Output
lblDom1L.Caption = intDomL1
lblDom1R.Caption = intDomR1

lblDom2L.Caption = intDomL2
lblDom2R.Caption = intDomR2

lblJoin.Caption = "Joinable: " & canJoin
End Sub

Private Sub Form_Load()
Randomize
End Sub

Public Function getDomino() As Integer
'Dec
Dim intRandom As Integer
'init
intRandom = 0
'Process
Do
intRandom = Int(67 * Rnd)
Loop While Right(intRandom, 1) > "6"
'Output
getDomino = intRandom
End Function

Public Function leftSide() As Integer
'Dec
Dim intDomL As Integer
'init
intDomL = 0
'process
intDomL = Left(getDomino, 1)
'output
leftSide = intDomL
End Function

Public Function rightSide() As Integer
'Dec
Dim intDomR As Integer
'init
intDomR = 0
'process
intDomR = Right(getDomino, 1)
'output
rightSide = intDomR
End Function

Public Function canJoin()
'Dec
Dim strCanJoin As String
'init
strCanJoin = ""
'process
If intDomL1 = intDomL2 Or intDomL1 = intDomR2 Or intDomR1 = intDomL2 Or intDomR1 = intDomR2 Then
    strCanJoin = "Yes"
Else
    strCanJoin = "No"
End If
'output
canJoin = strCanJoin
End Function
