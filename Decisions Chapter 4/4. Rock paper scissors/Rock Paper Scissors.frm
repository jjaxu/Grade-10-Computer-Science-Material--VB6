VERSION 5.00
Begin VB.Form RockPaperScissors 
   Caption         =   "Rock, Paper, Scissors"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   8400
      TabIndex        =   14
      Top             =   7080
      Width           =   3135
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
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
      Left            =   8400
      TabIndex        =   13
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Frame fraScore 
      Caption         =   "      You                 Computer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   8400
      TabIndex        =   12
      Top             =   1920
      Width           =   3135
      Begin VB.Label lblComputerScore 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1680
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblYourScore 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line3 
         X1              =   960
         X2              =   1800
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         X1              =   1560
         X2              =   1560
         Y1              =   120
         Y2              =   4440
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO!!!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2640
      TabIndex        =   4
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Frame fraChoice 
      Caption         =   "Pick your item"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
      Begin VB.OptionButton optScissors 
         Caption         =   "Scissors"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optPaper 
         Caption         =   "Paper"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optRock 
         Caption         =   "Rock"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label lblComputerChoice 
      Caption         =   "The Computer chose:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   15
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblScore 
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      TabIndex        =   11
      Top             =   360
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   8160
      X2              =   8160
      Y1              =   120
      Y2              =   7920
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   7575
   End
   Begin VB.Label lblComputerSide 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   9
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblVS 
      Alignment       =   2  'Center
      Caption         =   "VS."
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
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Image imgComputer 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Image imgYou 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblYou 
      Caption         =   "You"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblComputer 
      Caption         =   "Computer"
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
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Rock, Paper, Scissors"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "RockPaperScissors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: November 5 2013
'Purpose: Work with more decisions
Option Explicit
Dim intYourScore As Integer
Dim intComputerScore As Integer

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGO_Click()
'declare
Dim intComputer As Integer

'intitialize
intComputer = 0

'process
intComputer = Int(3 * Rnd) + 1

If intComputer = 1 Then
    lblComputerSide.Caption = "ROCK"
     imgComputer.Picture = LoadPicture(App.Path & "\Rock.jpg")
ElseIf intComputer = 2 Then
    lblComputerSide.Caption = "PAPER"
    imgComputer.Picture = LoadPicture(App.Path & "\Paper.jpg")
Else: lblComputerSide.Caption = "SCISSORS"
      imgComputer.Picture = LoadPicture(App.Path & "\Scissors.jpg")
End If

'Rock vs computer
If optRock.Value = True And intComputer = 1 Then
    lblStatus.Caption = "It's a Tie!"
ElseIf optRock.Value = True And intComputer = 2 Then
    lblStatus.Caption = "PAPER covers ROCK, the computer Wins!"
    intComputerScore = intComputerScore + 1
ElseIf optRock.Value = True And intComputer = 3 Then
    lblStatus.Caption = "ROCK smashes SCISSORS, you Win!"
    intYourScore = intYourScore + 1
End If

'Paper vs computer
If optPaper.Value = True And intComputer = 1 Then
    lblStatus.Caption = "PAPER covers ROCK, you Win!"
    intYourScore = intYourScore + 1
ElseIf optPaper.Value = True And intComputer = 2 Then
    lblStatus.Caption = "It's a Tie!"
ElseIf optPaper.Value = True And intComputer = 3 Then
    lblStatus.Caption = "SCISSORS cuts PAPER, the computer Wins!"
    intComputerScore = intComputerScore + 1
End If

'Scissors vs computer
If optScissors.Value = True And intComputer = 1 Then
    lblStatus.Caption = "ROCK smashes SCISSORS, the computer Wins!"
    intComputerScore = intComputerScore + 1
ElseIf optScissors.Value = True And intComputer = 2 Then
    lblStatus.Caption = "SCISSORS cuts PAPER, you Win!"
    intYourScore = intYourScore + 1
ElseIf optScissors.Value = True And intComputer = 3 Then
    lblStatus.Caption = "It's a Tie!"
End If

'output
lblYourScore.Caption = intYourScore
lblComputerScore.Caption = intComputerScore

End Sub

Private Sub cmdReset_Click()
intYourScore = 0
intComputerScore = 0
lblYourScore.Caption = intYourScore
lblComputerScore.Caption = intComputerScore

End Sub

Private Sub Form_Load()
Randomize
optRock.Value = True
End Sub

Private Sub optPaper_Click()
imgYou.Picture = LoadPicture(App.Path & "\Paper.jpg")
lblStatus.Caption = ""
lblComputerSide.Caption = ""
imgComputer.Picture = Nothing
End Sub

Private Sub optRock_Click()
imgYou.Picture = LoadPicture(App.Path & "\Rock.jpg")
lblStatus.Caption = ""
lblComputerSide.Caption = ""
imgComputer.Picture = Nothing
End Sub

Private Sub optScissors_Click()
imgYou.Picture = LoadPicture(App.Path & "\Scissors.jpg")
lblStatus.Caption = ""
lblComputerSide.Caption = ""
imgComputer.Picture = Nothing
End Sub
