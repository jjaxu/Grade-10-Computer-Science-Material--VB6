VERSION 5.00
Begin VB.Form frmTemp 
   Caption         =   "Temperature Converter"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   4800
      TabIndex        =   10
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtEnd 
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
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   4455
   End
   Begin VB.TextBox txtStart 
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
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   4455
   End
   Begin VB.OptionButton optCtoF 
      Caption         =   "Celsius to Fahrenheit"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.OptionButton optKtoF 
      Caption         =   "Kelvin to Fahrenheit"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   4095
   End
   Begin VB.OptionButton optCtoK 
      Caption         =   "Celsius to Kelvin"
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
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Frame fraType 
      Caption         =   "Choose converstion type:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
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
      TabIndex        =   1
      Top             =   7440
      Width           =   7935
   End
   Begin VB.Label lblEnding 
      Caption         =   "Ending Value:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label lblStarting 
      Caption         =   "Starting Value:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      Caption         =   "Temperature Converter"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: November 19 2013
'Purpose: To learn how to use decisions with loops
Option Explicit

Private Sub cmdConvert_Click()
'Declare
Dim sglStart As Single
Dim sglEnd As Single
Dim sglC As Single
Dim sglF As Single
Dim sglK As Single

'Initialize
sglStart = 0
sglEnd = 0
sglC = 0
sglF = 0
sglK = 0
'Input
sglStart = Val(txtStart.Text)
sglEnd = Val(txtEnd.Text)
'Process / Calculation / Output
lstOutput.Clear
If sglStart > sglEnd Then
    lstOutput.AddItem "Invalid Values!"
ElseIf optCtoF.Value = True Then
    Do While sglStart <= sglEnd
    sglC = sglStart
    sglF = 9 / 5 * sglC + 32
    lstOutput.AddItem sglC & "°C = " & sglF & "°F"
    sglStart = sglStart + 0.5
    Loop
ElseIf optKtoF.Value = True Then
    Do
    sglK = sglStart
    sglF = (sglK * 1.8) - 459.67
    lstOutput.AddItem sglK & "K = " & sglF & "°F"
    sglStart = sglStart + 0.5
    Loop While sglStart <= sglEnd
Else: Do
    sglC = sglStart
    sglK = sglC + 273.15
    lstOutput.AddItem sglC & "°C = " & sglK & "K"
    sglStart = sglStart + 0.5
    Loop While sglStart <= sglEnd
End If
End Sub

Private Sub Form_Load()
optCtoF.Value = True
End Sub

Private Sub optCtoF_Click()
lstOutput.Clear
End Sub

Private Sub optCtoK_Click()
lstOutput.Clear
End Sub

Private Sub optKtoF_Click()
lstOutput.Clear
End Sub

Private Sub txtEnd_Change()
lstOutput.Clear
End Sub

Private Sub txtStart_Change()
lstOutput.Clear
End Sub
