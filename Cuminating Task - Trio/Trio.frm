VERSION 5.00
Begin VB.Form frmTrio 
   Caption         =   "Trio"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   834
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   9480
      TabIndex        =   15
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Frame fraTurn 
      Caption         =   "Who goes first?"
      Height          =   855
      Left            =   5880
      TabIndex        =   7
      Top             =   6960
      Width           =   3375
      Begin VB.OptionButton optComp 
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optYou 
         Caption         =   "You"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game"
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
      Left            =   240
      TabIndex        =   6
      Top             =   6960
      Width           =   5415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   5
      Top             =   8040
      Width           =   3375
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Game"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   8040
      Width           =   5415
   End
   Begin VB.PictureBox picBoard 
      BackColor       =   &H8000000A&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000E&
      Height          =   5400
      Left            =   240
      ScaleHeight     =   356
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   596
      TabIndex        =   1
      Top             =   1320
      Width           =   9000
      Begin VB.Image imgComp3 
         Height          =   1560
         Left            =   7200
         Picture         =   "Trio.frx":0000
         Top             =   3720
         Width           =   1560
      End
      Begin VB.Image imgComp2 
         Height          =   1560
         Left            =   7200
         Picture         =   "Trio.frx":0164
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1560
      End
      Begin VB.Image imgComp1 
         Height          =   1560
         Left            =   7200
         Picture         =   "Trio.frx":02CB
         Top             =   120
         Width           =   1560
      End
      Begin VB.Image imgUser3 
         DragMode        =   1  'Automatic
         Height          =   1560
         Left            =   5520
         Picture         =   "Trio.frx":042F
         Top             =   3720
         Width           =   1560
      End
      Begin VB.Image imgUser2 
         DragMode        =   1  'Automatic
         Height          =   1560
         Left            =   5520
         Picture         =   "Trio.frx":0593
         Top             =   1920
         Width           =   1560
      End
      Begin VB.Image imgUser1 
         DragMode        =   1  'Automatic
         Height          =   1560
         Left            =   5520
         Picture         =   "Trio.frx":06F7
         Top             =   120
         Width           =   1560
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   360
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   240
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   120
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   360
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   360
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Label lblCompScore 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Left            =   11040
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblYouScore 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Left            =   9480
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblTitleCompScore 
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblTitleYourscore 
      Caption         =   "You"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblScore 
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblComp 
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblYou 
      Caption         =   "You"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Caption         =   "Trio"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmTrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Authors: Dustin Hu and Jackie Xu
'Date: January 13 2014
'Purpose: Culminating Task - Game of Tic Tac Toe from the things we learned
'http://titanpad.com/RNY98CSZsl
Option Explicit
'GV
Dim blnGameOver As Boolean
'Magic Square numbaaa
'User
Dim intUser1 As Integer
Dim intUser2 As Integer
Dim intUser3 As Integer

'Computer
Dim intComp1 As Integer
Dim intComp2 As Integer
Dim intComp3 As Integer

Private Sub cmdClear_Click()
'Scorekeep reset
lblYouScore.Caption = "0"
lblCompScore.Caption = "0"
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdReset_Click()
'User
intUser1 = -99
intUser2 = -98
intUser3 = -97

imgUser1.Left = 368
imgUser2.Left = 368
imgUser3.Left = 368

imgUser1.Top = 8
imgUser2.Top = 128
imgUser3.Top = 248
'Computer
intComp1 = -96
intComp2 = -95
intComp3 = -94

imgComp1.Left = 480
imgComp2.Left = 480
imgComp3.Left = 480

imgComp1.Top = 8
imgComp2.Top = 128
imgComp3.Top = 248
'Game Tracker
blnGameOver = False

fraTurn.Enabled = True
cmdStart.Enabled = True
optYou.Enabled = True
optComp.Enabled = True
picBoard.Enabled = False

End Sub

Private Sub cmdStart_Click()
picBoard.Enabled = True
fraTurn.Enabled = False
cmdStart.Enabled = False
optYou.Enabled = False
optComp.Enabled = False

If optComp.Value = True Then
    moveComputer intUser1, intUser2, intUser3, intComp1, intComp2, intComp3, imgComp1, imgComp2, imgComp3
End If
End Sub

Private Sub Form_Load()
Randomize
'User
intUser1 = -99
intUser2 = -98
intUser3 = -97
'Computer
intComp1 = -96
intComp2 = -95
intComp3 = -94
'Game Tracker
blnGameOver = False
optYou.Value = True
picBoard.Enabled = False
End Sub

Private Sub picBoard_DragDrop(Source As Control, X As Single, Y As Single)
'Declare
Dim intSquare As Integer
Dim intScoreYou As Integer
Dim intScoreComp As Integer

'Input
intScoreYou = lblYouScore.Caption
intScoreComp = lblCompScore.Caption

'checks if circle is placed in the correct spots
If getSquare(X, Y) <> 0 Then
    intSquare = getSquare(X, Y)
Else
    Exit Sub
End If


If X > 0 And X < 360 And Y > 0 And Y < 360 And IsSquareAvailable(intSquare, intUser1, intUser2, intUser3, _
intComp1, intComp2, intComp3) = True Then
    Source.Left = convertX(X)
    Source.Top = convertY(Y)
End If

'Validate the correct squares
If Source = imgUser1 Then
    If intUser1 = 0 Then
        Exit Sub
    Else
        intUser1 = getSquare(Source.Left, Source.Top)
    End If
        
ElseIf Source = imgUser2 Then
    If intUser2 = 0 Then
        Exit Sub
       
    Else
        intUser2 = getSquare(Source.Left, Source.Top)
    End If
       
ElseIf Source = imgUser3 Then
    If intUser3 = 0 Then
        Exit Sub
    Else
        intUser3 = getSquare(Source.Left, Source.Top)
    End If
End If

If intUser1 = 0 Then
    MsgBox ("Try Again")
    intUser1 = -99
    Exit Sub
ElseIf intUser2 = 0 Then
    MsgBox ("Try Again")
    intUser2 = -99
    Exit Sub
ElseIf intUser3 = 0 Then
    MsgBox ("Try Again")
    intUser3 = -99
    Exit Sub
End If


'The "AI" move sub
moveComputer intUser1, intUser2, intUser3, intComp1, intComp2, intComp3, imgComp1, imgComp2, imgComp3

If IsGameOver(intUser1, intUser2, intUser3) Then
    MsgBox ("You Win!")
    intScoreYou = intScoreYou + 1
    picBoard.Enabled = False
ElseIf IsGameOver(intComp1, intComp2, intComp3) Then
    MsgBox ("The Computer Wins!")
    intScoreComp = intScoreComp + 1
    picBoard.Enabled = False
End If


lblYouScore.Caption = intScoreYou
lblCompScore.Caption = intScoreComp

End Sub
Public Function getSquare(inputX As Single, inputY As Single) As Integer
'Column One
If inputX > 0 And inputX < 120 Then
   
    'Rows
    
    If inputY > 0 And inputY < 120 Then
        getSquare = 8
    ElseIf inputY > 120 And inputY < 240 Then
        getSquare = 3
    ElseIf inputY > 240 And inputY < 360 Then
        getSquare = 4
    Else
        getSquare = 0
    End If
'Column Two
ElseIf inputX > 120 And inputX < 240 Then
    'Rows
    
    If inputY > 0 And inputY < 120 Then
        getSquare = 1
    ElseIf inputY > 120 And inputY < 240 Then
        getSquare = 5
    ElseIf inputY > 240 And inputY < 360 Then
        getSquare = 9
    Else
        getSquare = 0
    End If
    
'Column Three
ElseIf inputX > 240 And inputX < 360 Then
    
    'Rows
    If inputY > 0 And inputY < 120 Then
        getSquare = 6
    ElseIf inputY > 120 And inputY < 240 Then
        getSquare = 7
    ElseIf inputY > 240 And inputY < 360 Then
        getSquare = 2
    Else
        getSquare = 0
    End If
End If
End Function

Public Function convertX(inputX) As Integer
If inputX > 0 And inputX < 120 Then
    convertX = 8
ElseIf inputX > 120 And inputX < 240 Then
    convertX = 128
ElseIf inputX > 240 And inputX < 360 Then
    convertX = 248
End If
End Function

Public Function convertY(inputY) As Integer
If inputY > 0 And inputY < 120 Then
    convertY = 8
ElseIf inputY > 120 And inputY < 240 Then
    convertY = 128
ElseIf inputY > 240 And inputY < 360 Then
    convertY = 248
End If
End Function

Public Function IsSquareAvailable(intSquare As Integer, intUser1 As Integer, intUser2 As Integer, intUser3 As Integer, _
intComp1 As Integer, intComp2 As Integer, intComp3 As Integer) As Boolean
'Checks validity of square, if it's not valid, it'll return false.
'Does so by checking each one individually.
'Maybe use a loop?
If intSquare = intUser1 Then
   IsSquareAvailable = False
ElseIf intSquare = intUser2 Then
    IsSquareAvailable = False
ElseIf intSquare = intUser3 Then
    IsSquareAvailable = False
ElseIf intSquare = intComp1 Then
    IsSquareAvailable = False
ElseIf intSquare = intComp2 Then
    IsSquareAvailable = False
ElseIf intSquare = intComp3 Then
    IsSquareAvailable = False
ElseIf intSquare = 0 Then
    IsSquareAvailable = False
Else
    IsSquareAvailable = True
End If

End Function

Public Function IsGameOver(piece1 As Integer, piece2 As Integer, piece3 As Integer) As Boolean
'Checks if values add up to 15
If piece1 + piece2 + piece3 = 15 Then
    IsGameOver = True
Else
    IsGameOver = False
End If
End Function

Public Function findRandomSquare(intUser1 As Integer, intUser2 As Integer, intUser3 As Integer, _
intComp1 As Integer, intComp2 As Integer, intComp3 As Integer) As Integer
'Iterates through all the possible points, looking for a possible one.
Dim intSquare As Integer

intSquare = Int(9 * Rnd) + 1
If IsSquareAvailable(intSquare, intUser1, intUser2, intUser3, intComp1, intComp2, intComp3) = False Then

    Do While IsSquareAvailable(intSquare, intUser1, intUser2, intUser3, intComp1, intComp2, intComp3) = False
        intSquare = Int(9 * Rnd + 1)
    Loop

End If
findRandomSquare = intSquare
End Function
Public Function getMarker(intComp1 As Integer, intComp2 As Integer, intComp3 As Integer) As Integer
If intComp1 < 0 Then
    getMarker = 1
ElseIf intComp2 < 0 Then
    getMarker = 2
ElseIf intComp3 < 0 Then
    getMarker = 3
Else
    getMarker = Int(3 * Rnd) + 1
End If
End Function

Public Sub putImage(intSquare As Integer, imgName As Object)
If intSquare = 8 Then
    imgName.Top = 8
    imgName.Left = 8
ElseIf intSquare = 1 Then
    imgName.Top = 8
    imgName.Left = 128
ElseIf intSquare = 6 Then
    imgName.Top = 8
    imgName.Left = 248
    
ElseIf intSquare = 3 Then
    imgName.Top = 128
    imgName.Left = 8
ElseIf intSquare = 5 Then
    imgName.Top = 128
    imgName.Left = 128
ElseIf intSquare = 7 Then
    imgName.Top = 128
    imgName.Left = 248
    
ElseIf intSquare = 4 Then
    imgName.Top = 248
    imgName.Left = 8
ElseIf intSquare = 9 Then
    imgName.Top = 248
    imgName.Left = 128
ElseIf intSquare = 2 Then
    imgName.Top = 248
    imgName.Left = 248
End If

End Sub

Public Sub moveComputer(intUser1 As Integer, intUser2 As Integer, intUser3 As Integer, intComp1 As Integer, intComp2 As Integer, intComp3 As Integer, _
imgComp1 As Object, imgComp2 As Object, imgComp3 As Object)
'This is the computer's turn, all contained within one little sub.
'Declare
Dim intMarker As Integer
Dim intSquare As Integer
Dim strAvailable As String

'Initialize
intMarker = 0
intSquare = 0
strAvailable = ""

'Input
intMarker = getMarker(intComp1, intComp2, intComp3)

strAvailable = getAvailableSquares(intUser1, intUser2, intUser3, intComp1, intComp2, intComp3)

'Process
intSquare = findPlayerTrio(intUser1, intUser2, intUser3, intComp1, intComp2, intComp2, strAvailable)
If intSquare = 0 Or intSquare = intUser1 Or intSquare = intUser2 Or intSquare = intUser3 _
Or intSquare = intComp1 Or intSquare = intComp2 Or intSquare = intComp3 Then
    intSquare = findRandomSquare(intUser1, intUser2, intUser3, intComp1, intComp2, intComp3)
End If

If intMarker = 1 Then
    putImage intSquare, imgComp1
    intComp1 = intSquare
ElseIf intMarker = 2 Then
    putImage intSquare, imgComp2
    intComp2 = intSquare
ElseIf intMarker = 3 Then
    putImage intSquare, imgComp3
    intComp3 = intSquare
End If
End Sub

Public Function getAvailableSquares(intUser1 As Integer, intUser2 As Integer, intUser3 As Integer, _
intComp1 As Integer, intComp2 As Integer, intComp3 As Integer) As String
' Checks out all the squares that are left - that is, ones that the computer can pick from.
'This should be parsed by the computer later on into a series of integers
'Which would then be fed into a list of available numbers.
'These would be the ones from which the computer picks from
'i.e, if getAvailableSquares is equal to "1,4,5,8,7,9", then the computer knows
'That numebrs 1, 4, 5, 8, 7, 9 are available
'And then, it would pass it through antoher function that checks CurrentSum of either Player or Computer
'If 15 - (number in getAvailableSquares) = CurrentSum, then it'll move the computer
'To that square.
'Else, it'll pick a random one.
Dim intCount
getAvailableSquares = "1,2,3,4,5,6,7,8,9"
For intCount = 1 To 9
If InStr(getAvailableSquares, intUser1) > 0 Then
    getAvailableSquares = Replace(getAvailableSquares, "," & intUser1, "")
ElseIf InStr(getAvailableSquares, intUser2) > 0 Then
    getAvailableSquares = Replace(getAvailableSquares, "," & intUser2, "")
ElseIf InStr(getAvailableSquares, intUser3) > 0 Then
    getAvailableSquares = Replace(getAvailableSquares, "," & intUser3, "")
ElseIf InStr(getAvailableSquares, intComp1) > 0 Then
    getAvailableSquares = Replace(getAvailableSquares, "," & intComp1, "")
ElseIf InStr(getAvailableSquares, intComp2) > 0 Then
    getAvailableSquares = Replace(getAvailableSquares, "," & intComp2, "")
ElseIf InStr(getAvailableSquares, intComp3) > 0 Then
    getAvailableSquares = Replace(getAvailableSquares, "," & intComp3, "")
End If
Next intCount
End Function

Public Function findPlayerTrio(intUser1 As Integer, intUser2 As Integer, intUser3 As Integer, intComp1 As Integer, intComp2 As Integer, _
intComp3 As Integer, strSquareAvailable As String) As Integer
'THis checks out which squares are still available and such
'Declare
Dim intPlayerSum As Integer
Dim intCount As Integer
Dim blnExit As Boolean
Dim intRound As Integer
Dim intCompSum As Integer

blnExit = False
intCount = 0
intRound = 1
intCompSum = 0
'The Exit is keeping track of whetehr or not intPlayerSum + intCount is equal to 15.
'If it equals to 15, then that would mean that it's the only square missing to win.
Do While blnExit = False
    If intRound = 1 Then
        intPlayerSum = intUser1 + intUser2
        intCompSum = intComp1 + intComp2
    ElseIf intRound = 2 Then
        intPlayerSum = intUser1 + intUser3
        intCompSum = intComp1 + intComp3
    Else
        intPlayerSum = intUser2 + intUser3
        intCompSum = intComp2 + intComp3
        blnExit = True
    End If
        
        
'This Seaches for possible computer moves
    For intCount = 1 To 9
        If InStr(strSquareAvailable, intCount) > 0 Then
            If intCompSum + intCount = 15 Then
                If Not (intUser1 = intCount Or intUser2 = intCount Or intUser3 = intCount) Then
                    blnExit = True
                    findPlayerTrio = intCount
                End If
            End If
        End If
    Next intCount
    
'This Seaches for possible player moves
    For intCount = 1 To 9
        If InStr(strSquareAvailable, intCount) > 0 Then
            If intPlayerSum + intCount = 15 Then
                If Not (intComp1 = intCount Or intComp2 = intCount Or intComp3 = intCount) Then
                    blnExit = True
                    findPlayerTrio = intCount
                End If
            End If
        End If
    Next intCount

    
    If intRound > 9 Then
        blnExit = True
        findPlayerTrio = 0
    End If
    intRound = intRound + 1
Loop

End Function

