VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Arrow Practice"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18495
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   18495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton down2 
      Caption         =   "Down alot"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton up2 
      Caption         =   "Up  alot"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton shoot 
      Caption         =   "SHOOT!!"
      BeginProperty Font 
         Name            =   "Roland"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton reset 
      Caption         =   "Play again?"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton down 
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton up 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Picture         =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin VB.Image arrow 
      Height          =   4410
      Left            =   15960
      Picture         =   "Form1.frx":05AC
      Top             =   3840
      Width           =   2700
   End
   Begin VB.Shape c0 
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   16800
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image bow2 
      Height          =   4410
      Left            =   2880
      Picture         =   "Form1.frx":08A6
      Top             =   360
      Width           =   2700
   End
   Begin VB.Image bow 
      Height          =   4410
      Left            =   2880
      Picture         =   "Form1.frx":0D6D
      Top             =   360
      Width           =   2700
   End
   Begin VB.Label lbl 
      Caption         =   $"Form1.frx":1319
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Shape c1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   16680
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   465
   End
   Begin VB.Shape c2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   16440
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   975
   End
   Begin VB.Shape c3 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1485
      Left            =   16200
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   1485
   End
   Begin VB.Shape c4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1995
      Left            =   15960
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Shape c5 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   15720
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Author: Jackie Xu'
'Date: September 21 2013'
Private Sub down_Click()
bow.Top = bow.Top + 25
If bow.Top >= 5040 Then down.Enabled = False
If bow.Top >= 4080 Then down2.Enabled = False
If bow.Top <= 5040 Then up.Enabled = True
If bow.Top >= 960 Then up2.Enabled = True
If bow.Top <= 4080 Then down2.Enabled = True
End Sub

Private Sub down2_Click()
bow.Top = bow.Top + 250
If bow.Top >= 4200 Then down2.Enabled = False
If bow.Top <= 4200 Then up2.Enabled = True
If bow.Top >= 120 Then up.Enabled = True
End Sub


Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
bow2.Visible = False
arrow.Visible = False
reset.Enabled = False
End Sub


Private Sub reset_Click()
bow.Visible = True
bow2.Visible = False
arrow.Visible = False
bow.Top = 360
up.Enabled = True
up2.Enabled = True
down.Enabled = True
down2.Enabled = True
reset.Enabled = False
shoot.Enabled = True
lbl.Caption = "Use the 'UP' and 'DOWN' buttons to control the arrow, press 'SHOOT!!'to fire the arrow. And 'Play again' to reset. 'EXIT' to quit."
End Sub

Private Sub shoot_Click()
bow.Visible = False
bow2.Top = bow.Top
bow2.Visible = True
arrow.Visible = True
arrow.Top = bow.Top + 120
If bow.Top >= 1320 And Number <= 3840 Then arrow.Left = 14520
If bow.Top <= 1320 Then arrow.Left = 16080
If bow.Top >= 3840 Then arrow.Left = 16080
If bow.Top <= 1320 Then lbl.Caption = "Missed!!"
If bow.Top >= 3840 Then lbl.Caption = "Missed!!"
If bow.Top >= 1320 And Number <= 3840 Then lbl.Caption = "Nice Shot!!!"
reset.Enabled = True
up.Enabled = False
up2.Enabled = False
down.Enabled = False
down2.Enabled = False
shoot.Enabled = False


End Sub

Private Sub up_Click()
bow.Top = bow.Top - 25
If bow.Top <= 120 Then up.Enabled = False
If bow.Top <= 960 Then up2.Enabled = False
If bow.Top >= 120 Then down.Enabled = True
If bow.Top <= 4020 Then down2.Enabled = True

End Sub

Private Sub up2_Click()
bow.Top = bow.Top - 250
If bow.Top <= 960 Then up2.Enabled = False
If bow.Top >= 960 Then down2.Enabled = True
End Sub
