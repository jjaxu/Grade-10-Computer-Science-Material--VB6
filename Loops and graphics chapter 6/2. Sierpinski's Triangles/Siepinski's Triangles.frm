VERSION 5.00
Begin VB.Form frmSierPT 
   Caption         =   "Form1"
   ClientHeight    =   13110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   874
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1006
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmV 
      Caption         =   "Number of Vertices"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   10560
      TabIndex        =   4
      Top             =   2880
      Width           =   4215
      Begin VB.OptionButton Option3 
         Caption         =   "5 Vertices"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "4 Vertices"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 Vertices (Default)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   120
      TabIndex        =   2
      Top             =   12000
      Width           =   10095
   End
   Begin VB.PictureBox picDraw 
      Height          =   10215
      Left            =   240
      ScaleHeight     =   677
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   661
      TabIndex        =   1
      Top             =   1560
      Width           =   9975
   End
   Begin VB.Label lblExtra 
      Caption         =   "Extra Options"
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
      Left            =   10560
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label lblTitle 
      Caption         =   "Sierpinski's Triangles"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
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
      Width           =   6615
   End
End
Attribute VB_Name = "frmSierPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: December 6 2013
'Purpose: working with random numbers as well as graphics such as PSet
Option Explicit

'GV
Dim intX1 As Integer
Dim intY1 As Integer

Dim intX2 As Integer
Dim intY2 As Integer

Dim intX3 As Integer
Dim intY3 As Integer

Dim intX4 As Integer
Dim intY4 As Integer

Dim intCounter As Integer
Dim intPointCount As Integer

Dim intRandom As Integer

Private Sub cmdClear_Click()
picDraw.Cls
intPointCount = 0
End Sub

Private Sub Form_Load()
Randomize
'Initialize
intCounter = 0
intPointCount = 0
intRandom = 0
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)




'Input / NO INPUT
'Process / Output
If intPointCount <= 4 Then
    intPointCount = intPointCount + 1
    picDraw.PSet (X, Y), vbBlue

    If intPointCount = 1 Then
        intX1 = X
        intY1 = Y
    ElseIf intPointCount = 2 Then
        intX2 = X
        intY2 = Y
    ElseIf intPointCount = 3 Then
        intX3 = X
        intY3 = Y
    Else
        intX4 = X
        intY4 = Y
    End If

        If intPointCount = 4 Then
            For intCounter = 1 To 32000
                intRandom = Int(Rnd * 3) + 1
                If intRandom = 1 Then
                    intX4 = (intX4 + intX1) / 2
                    intY4 = (intY4 + intY1) / 2
                    picDraw.PSet (intX4, intY4), vbBlue
                ElseIf intRandom = 2 Then
                    intX4 = (intX4 + intX2) / 2
                    intY4 = (intY4 + intY2) / 2
                    picDraw.PSet (intX4, intY4), vbRed
                Else
                    intX4 = (intX4 + intX3) / 2
                    intY4 = (intY4 + intY3) / 2
                    picDraw.PSet (intX4, intY4), vbGreen
                End If
  
    

    
            Next intCounter
    End If
End If
End Sub
