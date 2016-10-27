VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Light Bulb"
   ClientHeight    =   10065
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16290
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   671
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1086
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdoff 
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6480
      TabIndex        =   2
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton cmdon 
      Caption         =   "ON"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6480
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label LBLmessage 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   6000
      Width           =   8655
   End
   Begin VB.Image imgdisplay 
      Height          =   5415
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By:     Jackie'
'Date: september 18 2013'


Private Sub Command1_Click()

End Sub

Private Sub cmdoff_Click()
imgdisplay.Picture = LoadPicture(App.Path & "\bulboff.jpg")
cmdoff.Enabled = False
cmdon.Enabled = True
LBLmessage = "The light is now OFF"
End Sub

Private Sub cmdon_Click()
imgdisplay.Picture = LoadPicture(App.Path & "\bulbon.jpg")
cmdon.Enabled = Flase
cmdoff.Enabled = True
LBLmessage = "The light is now ON"
End Sub

Private Sub Form_Load()
imgdisplay.Picture = LoadPicture(App.Path & "\bulbon.jpg")
cmdon.Enabled = Flase
cmdoff.Enabled = True
LBLmessage.Caption = ""
End Sub

Private Sub imgdisplay_Click()
LBLmessage = "lol"
End Sub

Private Sub optbulboff_Click()
imgdisplay.Picture = LoadPicture(App.Path & "\bulboff.jpg")
LBLmessage = "The light is now OFF"
End Sub

Private Sub optbulbon_Click()
imgdisplay.Picture = LoadPicture(App.Path & "\bulbon.jpg")
LBLmessage = "The light is now ON"
End Sub
