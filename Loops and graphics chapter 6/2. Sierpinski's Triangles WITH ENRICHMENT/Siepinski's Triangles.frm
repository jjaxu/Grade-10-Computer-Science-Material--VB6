VERSION 5.00
Begin VB.Form frmSierPT 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   729
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1006
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPercent 
      Alignment       =   2  'Center
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
      Left            =   10560
      TabIndex        =   9
      Text            =   "50"
      Top             =   7920
      Width           =   4215
   End
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
      Begin VB.OptionButton opt5V 
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
      Begin VB.OptionButton opt4V 
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
      Begin VB.OptionButton opt3V 
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
   Begin VB.Label lblPercent 
      Caption         =   "Percentage of approach (%)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10560
      TabIndex        =   8
      Top             =   6480
      Width           =   4215
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

Dim intX5 As Integer
Dim intY5 As Integer

Dim intTravX As Integer
Dim intTravY As Integer

Dim lngCounter As Long
Dim intPointCount As Integer

Dim intRandom As Integer

Dim sglPercent1 As Single
Dim sglPercent2 As Single

Private Sub cmdClear_Click()
picDraw.Cls
intPointCount = 0
lngCounter = 0
intRandom = 0
sglPercent1 = 0
sglPercent2 = 0
End Sub

Private Sub Form_Load()
Randomize
opt3V.Value = True
'Initialize
lngCounter = 0
intPointCount = 0
intRandom = 0
sglPercent1 = 0
sglPercent2 = 0
End Sub

Private Sub opt3V_Click()
picDraw.Cls
intPointCount = 0
End Sub

Private Sub opt4V_Click()
picDraw.Cls
intPointCount = 0
End Sub

Private Sub opt5V_Click()
picDraw.Cls
intPointCount = 0
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Input
sglPercent1 = Val(txtPercent.Text) / 100

'Process / Output
sglPercent2 = 1 - sglPercent1

If intPointCount <= 4 And opt3V.Value = True Then
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
        intTravX = X
        intTravY = Y
    End If

        If intPointCount = 4 Then
            For lngCounter = 1 To 50000
                intRandom = Int(Rnd * 3) + 1
                If intRandom = 1 Then
                    intTravX = (intTravX * sglPercent1 + intX1 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY1 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbBlue
                ElseIf intRandom = 2 Then
                    intTravX = (intTravX * sglPercent1 + intX2 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY2 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbRed
                Else
                    intTravX = (intTravX * sglPercent1 + intX3 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY3 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbGreen
                End If
            Next lngCounter
        End If

'4 Vertices
ElseIf intPointCount <= 5 And opt4V.Value = True Then
   
    
 
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
    ElseIf intPointCount = 4 Then
        intX4 = X
        intY4 = Y
    Else
        intTravX = X
        intTravY = Y
    End If

        If intPointCount = 5 Then
            For lngCounter = 1 To 50000
                intRandom = Int(Rnd * 4) + 1
                If intRandom = 1 Then
                    intTravX = (intTravX * sglPercent1 + intX1 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY1 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbBlue
                ElseIf intRandom = 2 Then
                    intTravX = (intTravX * sglPercent1 + intX2 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY2 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbRed
                ElseIf intRandom = 3 Then
                    intTravX = (intTravX * sglPercent1 + intX3 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY3 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbGreen
                Else
                    intTravX = (intTravX * sglPercent1 + intX4 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY4 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbCyan
                End If
            Next lngCounter
        End If
'5 vertices
ElseIf intPointCount <= 6 And opt5V.Value = True Then
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
    ElseIf intPointCount = 4 Then
        intX4 = X
        intY4 = Y
    ElseIf intPointCount = 5 Then
        intX5 = X
        intY5 = Y
    Else
        intTravX = X
        intTravY = Y
    End If

        If intPointCount = 6 Then
            For lngCounter = 1 To 50000
                intRandom = Int(Rnd * 5) + 1
                If intRandom = 1 Then
                    intTravX = (intTravX * sglPercent1 + intX1 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY1 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbBlue
                ElseIf intRandom = 2 Then
                    intTravX = (intTravX * sglPercent1 + intX2 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY2 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbRed
                ElseIf intRandom = 3 Then
                    intTravX = (intTravX * sglPercent1 + intX3 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY3 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbGreen
                ElseIf intRandom = 4 Then
                    intTravX = (intTravX * sglPercent1 + intX4 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY4 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbCyan
                Else
                    intTravX = (intTravX * sglPercent1 + intX5 * sglPercent2)
                    intTravY = (intTravY * sglPercent1 + intY5 * sglPercent2)
                    picDraw.PSet (intTravX, intTravY), vbMagenta
                End If
            Next lngCounter
        End If
End If
End Sub

Private Sub txtPercent_Change()
picDraw.Cls
intPointCount = 0
End Sub
