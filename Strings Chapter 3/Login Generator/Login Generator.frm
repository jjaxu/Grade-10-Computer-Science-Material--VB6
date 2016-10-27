VERSION 5.00
Begin VB.Form LoginGenerator 
   Caption         =   "Login Generator"
   ClientHeight    =   11490
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLL 
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
      Left            =   3120
      TabIndex        =   14
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtLF 
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
      Left            =   3120
      TabIndex        =   11
      Top             =   2880
      Width           =   2175
   End
   Begin VB.OptionButton opt4thD 
      Caption         =   "Disable custom login"
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
      Left            =   6120
      TabIndex        =   10
      Top             =   2280
      Width           =   2535
   End
   Begin VB.OptionButton opt4thE 
      Caption         =   "Enable custom login"
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
      TabIndex        =   9
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "GENERATE!!!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   8415
   End
   Begin VB.TextBox txtInput 
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
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   8415
   End
   Begin VB.Label lblL 
      Caption         =   "Number of letters in last name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label lblF 
      Caption         =   "Number of letters in first name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblLogin4 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   10560
      Width           =   8415
   End
   Begin VB.Label lblLogin4M 
      Caption         =   "Custom Login:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   9960
      Width           =   8415
   End
   Begin VB.Label lblLogin3M 
      Caption         =   "Login 3:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   8400
      Width           =   8415
   End
   Begin VB.Label llblLogin2M 
      Caption         =   "Login 2:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   6960
      Width           =   8415
   End
   Begin VB.Label lblLogin1M 
      Caption         =   "Login 1:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5520
      Width           =   8415
   End
   Begin VB.Label lblLogin1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   6120
      Width           =   8415
   End
   Begin VB.Label lblLogin2 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   7560
      Width           =   8415
   End
   Begin VB.Label lblLogin3 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   9000
      Width           =   8415
   End
   Begin VB.Label lblMessage 
      Caption         =   "Please enter your first and last name separated by a space below, then click ""Generate"" to generate your logins."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "LoginGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: october 18 2013
'purpose: Learing how to cut string varaibles into pieces
Option Explicit

Private Sub cmdGenerate_Click()
'Declare
Dim strInput As String
Dim strLogin1 As String
Dim strLogin2 As String
Dim strLogin3 As String
Dim strLogin4 As String
Dim strFirst As String
Dim strLast As String
Dim intFirst As Integer
Dim intLast As Integer

'initialize
strInput = ""
strLogin1 = ""
strLogin2 = ""
strLogin3 = ""
strLogin4 = ""
strFirst = ""
strLast = ""
intFirst = 0
intLast = 0
'input

strInput = txtInput.Text
intFirst = Val(txtLF.Text)
intLast = Val(txtLL.Text)

'calculation/process
Trim (strInput)
strFirst = Left(strInput, InStr(strInput, " "))
strFirst = Trim(strFirst)

strLast = Mid(strInput, InStr(strInput, " "))
strLast = Trim(strLast)

strLogin1 = Left(strFirst, 4) & Left(strLast, 3)
strLogin2 = Mid(strLast, 2) & Left(strFirst, 1)
strLogin3 = Left(strFirst, 1) & Right(strLast, 4)
strLogin4 = Left(strFirst, intFirst) & Left(strLast, intLast)
'output
lblLogin1.Caption = strLogin1
lblLogin2.Caption = strLogin2
lblLogin3.Caption = strLogin3
lblLogin4.Caption = strLogin4
End Sub

Private Sub Form_Load()
opt4thD.Value = True
cmdGenerate.Enabled = False
End Sub

Private Sub opt4thD_Click()
lblLogin4.Visible = False
lblLogin4M.Visible = False
lblF.Visible = False
lblL.Visible = False
txtLF.Visible = False
txtLL.Visible = False
End Sub

Private Sub opt4thE_Click()
lblLogin4.Visible = True
lblLogin4M.Visible = True
lblF.Visible = True
lblL.Visible = True
txtLF.Visible = True
txtLL.Visible = True
End Sub

Private Sub txtInput_Change()
Dim strInput As String
strInput = ""
strInput = txtInput.Text
lblLogin1.Caption = ""
lblLogin2.Caption = ""
lblLogin3.Caption = ""
lblLogin4.Caption = ""
If txtInput.Text = "" Then cmdGenerate.Enabled = False
If txtInput.Text <> "" Then cmdGenerate.Enabled = True
If InStr(strInput, " ") = 0 Then cmdGenerate.Enabled = False
If InStr(strInput, " ") <> 0 Then cmdGenerate.Enabled = True
End Sub
