VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Picture Changer"
   ClientHeight    =   12930
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5520
      TabIndex        =   5
      Top             =   11640
      Width           =   4935
   End
   Begin VB.CommandButton clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   11640
      Width           =   4935
   End
   Begin VB.CommandButton violin 
      Caption         =   "Violin"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      TabIndex        =   3
      Top             =   10560
      Width           =   2295
   End
   Begin VB.CommandButton flag 
      Caption         =   "Canadian Flag"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   10560
      Width           =   2295
   End
   Begin VB.CommandButton earth 
      Caption         =   "Earth"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   10560
      Width           =   2295
   End
   Begin VB.CommandButton cube 
      Caption         =   "Rubik's Cube"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   10560
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Select any picture you want to see, press ""Clear"" to clear the screen, and press ""Exit"" to quit."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   10095
   End
   Begin VB.Image img 
      Height          =   9975
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu'
'Date: September 21 2013'

Private Sub clear_Click()
img.Visible = False
cube.Enabled = True
earth.Enabled = True
flag.Enabled = True
violin.Enabled = True
lbl.Visible = True
clear.Enabled = False
End Sub

Private Sub cube_Click()
img.Picture = LoadPicture(App.Path & "\cube.jpg")
img.Visible = True
cube.Enabled = False
earth.Enabled = True
flag.Enabled = True
violin.Enabled = True
lbl.Visible = False
clear.Enabled = True
End Sub

Private Sub earth_Click()
img.Picture = LoadPicture(App.Path & "\earth.jpg")
img.Visible = True
earth.Enabled = False
cube.Enabled = True
flag.Enabled = True
clear.Enabled = True
violin.Enabled = True
lbl.Visible = False
clear.Enabled = True
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub flag_Click()
img.Picture = LoadPicture(App.Path & "\flag.jpg")
img.Visible = True
flag.Enabled = False
cube.Enabled = True
violin.Enabled = True
earth.Enabled = True
lbl.Visible = False
clear.Enabled = True
End Sub

Private Sub Form_Load()
clear.Enabled = False
End Sub


Private Sub violin_Click()
img.Picture = LoadPicture(App.Path & "\violin.jpg")
img.Visible = True
violin.Enabled = False
cube.Enabled = True
earth.Enabled = True
clear.Enabled = True
flag.Enabled = True
lbl.Visible = False
End Sub
