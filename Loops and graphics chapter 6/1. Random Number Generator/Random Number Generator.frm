VERSION 5.00
Begin VB.Form frmRdnGen 
   Caption         =   "Random Number Generator"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   632
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5505
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10695
   End
   Begin VB.Label lbl225 
      Caption         =   " 225"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lbl200 
      Caption         =   " 200"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lbl175 
      Caption         =   " 175"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lbl50 
      Caption         =   " 150"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label lbl125 
      Caption         =   " 125"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label lbl100 
      Caption         =   " 100"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lbl75 
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label50 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label25 
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label lbl250 
      Caption         =   " 250"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lbl0 
      Caption         =   "  0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label lblTerm 
      Caption         =   "# Times"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblSum 
      Caption         =   "Sum"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label lblSumNumbers 
      Caption         =   "  2    3     4    5    6    7     8     9   10  11  12"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   8160
      Width           =   6615
   End
   Begin VB.Shape shp12 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   10200
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp11 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   9600
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp10 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   9000
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp9 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   8400
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp8 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   7800
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp7 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   7200
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp6 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   6600
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp5 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   6000
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp4 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   5400
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   4800
      Top             =   8040
      Width           =   615
   End
   Begin VB.Shape shp2 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   4200
      Top             =   8040
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   280
      X2              =   280
      Y1              =   216
      Y2              =   536
   End
   Begin VB.Line Line1 
      X1              =   280
      X2              =   720
      Y1              =   536
      Y2              =   536
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Random Number Generator XT-1000"
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
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "frmRdnGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: December 2 2013
'Purpose: Explore for loops with random numbers
Option Explicit

Private Sub cmdGenerate_Click()
'Declare
Dim intNumber1 As Integer
Dim intNumber2 As Integer
Dim intCounter As Integer
Dim intAnswer As Integer

Dim int2Count As Integer
Dim int3Count As Integer
Dim int4Count As Integer
Dim int5Count As Integer
Dim int6Count As Integer
Dim int7Count As Integer
Dim int8Count As Integer
Dim int9Count As Integer
Dim int10Count As Integer
Dim int11Count As Integer
Dim int12Count As Integer
'Initialize
intNumber1 = 0
intNumber2 = 0
intCounter = 0
intAnswer = 0

int2Count = 0
int3Count = 0
int4Count = 0
int5Count = 0
int6Count = 0
int7Count = 0
int8Count = 0
int9Count = 0
int10Count = 0
int11Count = 0
int12Count = 0
'NO Input
'Process / Calculation
lstOutput.Clear
For intCounter = 1 To 1000
    intNumber1 = Int(6 * Rnd) + 1
    intNumber2 = Int(6 * Rnd) + 1
    intAnswer = intNumber1 + intNumber2
    If intAnswer = 2 Then
        int2Count = int2Count + 1
        
        
    ElseIf intAnswer = 3 Then
        int3Count = int3Count + 1
        
    ElseIf intAnswer = 4 Then
        int4Count = int4Count + 1
        
    ElseIf intAnswer = 5 Then
        int5Count = int5Count + 1
    
    ElseIf intAnswer = 6 Then
        int6Count = int6Count + 1
    
    ElseIf intAnswer = 7 Then
        int7Count = int7Count + 1
    
    ElseIf intAnswer = 8 Then
        int8Count = int8Count + 1
    
    ElseIf intAnswer = 9 Then
        int9Count = int9Count + 1
        
    ElseIf intAnswer = 10 Then
        int10Count = int10Count + 1
        
    ElseIf intAnswer = 11 Then
        int11Count = int11Count + 1
        
    Else: int12Count = int12Count + 1
End If
    
Next intCounter
'Output
lstOutput.AddItem "Sum is 2: " & int2Count & " times"
lstOutput.AddItem "Sum is 3: " & int3Count & " times"
lstOutput.AddItem "Sum is 4: " & int4Count & " times"
lstOutput.AddItem "Sum is 5: " & int5Count & " times"
lstOutput.AddItem "Sum is 6: " & int6Count & " times"
lstOutput.AddItem "Sum is 7: " & int7Count & " times"
lstOutput.AddItem "Sum is 8: " & int8Count & " times"
lstOutput.AddItem "Sum is 9: " & int9Count & " times"
lstOutput.AddItem "Sum is 10: " & int10Count & " times"
lstOutput.AddItem "Sum is 11: " & int11Count & " times"
lstOutput.AddItem "Sum is 12: " & int12Count & " times"

shp2.Height = int2Count + 9
shp2.Top = (250 - int2Count) + 278

shp3.Height = int3Count + 9
shp3.Top = (250 - int3Count) + 278

shp4.Height = int4Count + 9
shp4.Top = (250 - int4Count) + 278

shp5.Height = int5Count + 9
shp5.Top = (250 - int5Count) + 278

shp6.Height = int6Count + 9
shp6.Top = (250 - int6Count) + 278

shp7.Height = int7Count + 9
shp7.Top = (250 - int7Count) + 278

shp8.Height = int8Count + 9
shp8.Top = (250 - int8Count) + 278

shp9.Height = int9Count + 9
shp9.Top = (250 - int9Count) + 278

shp10.Height = int10Count + 9
shp10.Top = (250 - int10Count) + 278

shp11.Height = int11Count + 9
shp11.Top = (250 - int11Count) + 278

shp12.Height = int12Count + 9
shp12.Top = (250 - int12Count) + 278



End Sub

Private Sub Form_Load()
Randomize
End Sub

