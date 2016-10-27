VERSION 5.00
Begin VB.Form VolumeAreaCalculator 
   Caption         =   "Volume and Surface Area"
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   13
      Top             =   5760
      Width           =   10335
   End
   Begin VB.PictureBox picSA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1275
      ScaleWidth      =   10275
      TabIndex        =   12
      Top             =   8520
      Width           =   10335
   End
   Begin VB.PictureBox picVolume 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1395
      ScaleWidth      =   10275
      TabIndex        =   11
      Top             =   6840
      Width           =   10335
   End
   Begin VB.TextBox txtHeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   10
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox txtWidth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   9
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox txtLength 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3840
      TabIndex        =   8
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Frame calculation 
      Caption         =   "Polyherdron"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optSphere 
         Caption         =   "Sphere"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optCylinder 
         Caption         =   "Cylinder"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optRectangle 
         Caption         =   "Rectangular Prism"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image imgDisplay 
      Height          =   3165
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3480
   End
   Begin VB.Label lblHeight 
      Caption         =   "Enter Height:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label lblWidth 
      Caption         =   "Enter Width:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label lblLength 
      Caption         =   "Enter Length:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lbltitle 
      Caption         =   "Volume and Surface Area calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "VolumeAreaCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu'
'Date: September 30 2013'
'Purpose: Learn how to use commands'
Option Explicit

Private Sub cmdCalculate_Click()
'delare
Dim SglLength As Single
Dim SglWidth As Single
Dim SglHeight As Single
Dim SglSA As Single
Dim SglVolume As Single
'initialize
SglLength = 0
SglWidth = 0
SglHeight = 0
SglSA = 0
SglVolume = 0
'input
SglLength = Val(txtLength.Text)
SglWidth = Val(txtWidth.Text)
SglHeight = Val(txtHeight.Text)
'process/calculation
If optRectangle.Value = True And SglLength > 0 Then SglVolume = SglLength * SglWidth * SglHeight
If optRectangle.Value = True Then SglSA = 2 * (SglLength * SglWidth + SglLength * SglHeight + SglWidth * SglHeight)
If optCylinder.Value = True Then SglVolume = SglWidth ^ 2 * 3.14159265358979 * SglHeight
If optCylinder.Value = True Then SglSA = 2 * (SglWidth ^ 2 * 3.14159265358979) + SglHeight * 2 * SglWidth * 3.14159265358979
If optSphere.Value = True Then SglVolume = (4 / 3) * 3.14159265358979 * SglWidth ^ 3
If optSphere.Value = True Then SglSA = 4 * SglWidth ^ 2 * 3.14159265358979
'output
picVolume.Cls
picSA.Cls
If 99999999 > SglLength And SglWidth And SglHeight > 0 Then picVolume.Print "The Volume is " & SglVolume & " Cubic Units."
If 99999999 > SglLength And SglWidth And SglHeight > 0 Then picSA.Print "The Surface Area is " & SglSA & " Square Units."

If SglLength < 0 Then picVolume.Cls
If SglLength < 0 Then picSA.Cls
If SglLength < 0 Then picVolume.Print "Please enter a positive value."

If SglWidth < 0 Then picVolume.Cls
If SglWidth < 0 Then picSA.Cls
If SglWidth < 0 Then picVolume.Print "Please enter a positive value."

If SglHeight < 0 Then picVolume.Cls
If SglHeight < 0 Then picSA.Cls
If SglHeight < 0 Then picVolume.Print "Please enter a positive value."
End Sub

Private Sub Form_Load()
imgDisplay.Picture = LoadPicture(App.Path & "\rectangle.bmp")
End Sub

Private Sub optCylinder_Click()
imgDisplay.Picture = LoadPicture(App.Path & "\cylinder.bmp")
txtLength.Visible = False
lblLength.Visible = False
txtHeight.Visible = True
lblHeight.Visible = True

lblWidth.Caption = "Enter Width"
End Sub

Private Sub optRectangle_Click()
imgDisplay.Picture = LoadPicture(App.Path & "\rectangle.bmp")
txtLength.Visible = True
lblLength.Visible = True
txtWidth.Visible = True
lblWidth.Visible = True
txtHeight.Visible = True
lblHeight.Visible = True
lblWidth.Caption = "Enter Width"
End Sub

Private Sub optSphere_Click()
imgDisplay.Picture = LoadPicture(App.Path & "\sphere.bmp")
txtLength.Visible = False
lblLength.Visible = False
txtHeight.Visible = False
lblHeight.Visible = False
lblWidth.Caption = "Enter Radius"
End Sub
