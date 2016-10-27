VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Light Bulb"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9240
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
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optbulbon 
      Caption         =   "Light ON"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6120
      TabIndex        =   1
      Top             =   3960
      Width           =   2655
   End
   Begin VB.OptionButton optbulboff 
      Caption         =   "Light OFF"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
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
      TabIndex        =   2
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


Private Sub Form_Load()
optbulbon.Value = True
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
