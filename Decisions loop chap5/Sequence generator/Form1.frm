VERSION 5.00
Begin VB.Form frmSequence 
   Caption         =   "Sequence Generator"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   8775
   End
   Begin VB.OptionButton optF 
      Caption         =   "Fibonacci Sequence"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.OptionButton optG 
      Caption         =   "Geometric Sequence"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   2895
   End
   Begin VB.OptionButton optA 
      Caption         =   "Arithmetic Sequence"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   8775
   End
   Begin VB.TextBox txtNumbT 
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
      Left            =   5280
      TabIndex        =   6
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox txtCommD 
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
      Left            =   5280
      TabIndex        =   5
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox txtStart 
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
      Left            =   5280
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label lblNumbT 
      Caption         =   "Number of Terms:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label lblCommD 
      Caption         =   "Common Difference:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label lblStart 
      Caption         =   "Starting Number:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      Caption         =   "Sequence Generator"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
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
      Width           =   5055
   End
End
Attribute VB_Name = "frmSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: November 21 2013
'Purpose: To play around with more decision loops
Option Explicit

Private Sub cmdGenerate_Click()
'Dim
Dim sglStart As Single
Dim sglCommD As Single
Dim intNumbT As Integer
Dim sglResult As Single
Dim sglCounter As Single
Dim sglCurrent As Single
Dim sglPrevious As Single
'Initialize
sglStart = 0
sglCommD = 0
intNumbT = 0
sglResult = 0
sglCounter = 0
sglCurrent = 0
sglPrevious = 0
'Input
sglStart = Val(txtStart.Text)
sglCommD = Val(txtCommD.Text)
intNumbT = Val(txtNumbT.Text)
'Process / Output
lstOutput.Clear
If optA.Value = True And intNumbT = 0 Then
    lstOutput.Clear
ElseIf optA.Value = True And intNumbT = 1 Then
    lstOutput.AddItem sglStart
ElseIf optA.Value = True And intNumbT = 2 Then
    lstOutput.AddItem sglStart
    lstOutput.AddItem sglStart + sglCommD

ElseIf optA.Value = True Then
    lstOutput.AddItem sglStart
    sglResult = sglStart + sglCommD
    lstOutput.AddItem sglResult
    sglCounter = sglCounter + 3
        Do
            sglResult = sglResult + sglCommD
            sglCounter = sglCounter + 1
            lstOutput.AddItem sglResult
        Loop While sglCounter <= intNumbT

ElseIf optG.Value = True And intNumbT = 0 Then
    lstOutput.Clear
ElseIf optG.Value = True And intNumbT = 1 Then
    lstOutput.AddItem sglStart
ElseIf optG.Value = True And intNumbT = 2 Then
    lstOutput.AddItem sglStart
    lstOutput.AddItem sglStart * sglCommD




ElseIf optG.Value = True Then
    lstOutput.AddItem sglStart
    sglResult = sglStart * sglCommD
    lstOutput.AddItem sglResult
    sglCounter = sglCounter + 3
        Do
            sglResult = sglResult * sglCommD
            sglCounter = sglCounter + 1
            lstOutput.AddItem sglResult
        Loop While sglCounter <= intNumbT
    
Else
    If intNumbT = 0 Then
        lstOutput.Clear
    ElseIf intNumbT = 1 Then
        lstOutput.AddItem "0"
    ElseIf intNumbT = 2 Then
        lstOutput.AddItem "0"
        lstOutput.AddItem "1"
    Else
        lstOutput.AddItem "0"
        lstOutput.AddItem "1"
        sglCounter = sglCounter + 3
    
        sglPrevious = 0
        sglCurrent = 1
        Do
            sglResult = sglCurrent + sglPrevious
            sglPrevious = sglCurrent
            sglCurrent = sglResult
            sglCounter = sglCounter + 1
            lstOutput.AddItem sglCurrent
        Loop While sglCounter <= intNumbT
        End If
    End If


End Sub

Private Sub Form_Load()
optA.Value = True
End Sub

Private Sub optA_Click()
txtCommD.Enabled = True
txtStart.Enabled = True
lblCommD.Enabled = True
lblStart.Enabled = True
lstOutput.Clear
End Sub

Private Sub optF_Click()
txtCommD.Enabled = False
txtStart.Enabled = False
lblCommD.Enabled = False
lblStart.Enabled = False
txtCommD.Text = ""
txtStart.Text = ""
lstOutput.Clear
End Sub

Private Sub optG_Click()
txtCommD.Enabled = True
txtStart.Enabled = True
lblCommD.Enabled = True
lblStart.Enabled = True
lstOutput.Clear
End Sub

Private Sub txtCommD_Change()
lstOutput.Clear
End Sub

Private Sub txtNumbT_Change()
lstOutput.Clear
End Sub

Private Sub txtStart_Change()
lstOutput.Clear
End Sub
