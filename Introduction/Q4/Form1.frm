VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optdisagree 
      BackColor       =   &H80000009&
      Caption         =   "DISAGREE"
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   5760
      Width           =   2535
   End
   Begin VB.OptionButton optagree 
      BackColor       =   &H80000009&
      Caption         =   "AGREE"
      Height          =   495
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtuser 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdcontinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   6360
      Width           =   3255
   End
   Begin VB.CommandButton cmdlogout 
      Caption         =   "Logout"
      Height          =   735
      Left            =   7920
      TabIndex        =   6
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Image imgsignin 
      Height          =   330
      Left            =   4200
      Picture         =   "Form1.frx":0000
      Top             =   4200
      Width           =   1050
   End
   Begin VB.Image imgrbc 
      Height          =   8160
      Left            =   0
      Picture         =   "Form1.frx":127A
      Top             =   0
      Width           =   11790
   End
   Begin VB.Image imgagree 
      Height          =   7155
      Left            =   0
      Picture         =   "Form1.frx":2C917
      Top             =   0
      Width           =   11280
   End
   Begin VB.Label lblty 
      Caption         =   "Thank you for logging in"
      BeginProperty Font 
         Name            =   "OCR A Std"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcontinue_Click()
optagree.Visible = False
optdisagree.Visible = False
cmdcontinue.Visible = False
imgagree.Visible = False
cmdlogout.Visible = True
End Sub

Private Sub cmdlogout_Click()
imgrbc.Visible = True
txtuser.Visible = True
txtpass.Visible = True
imgsignin.Visible = True
imgagree.Visible = True
txtuser.Text = ""
txtpass.Text = ""
cmdlogout.Visible = False
imgsignin.Enabled = False
End Sub

Private Sub Form_Load()
imgsignin.Enabled = False
optagree.Visible = False
optdisagree.Visible = False
cmdcontinue.Visible = False
cmdlogout.Visible = False
End Sub

Private Sub imgrbc_Click()

End Sub

Private Sub imgsignin_Click()
imgrbc.Visible = False
txtuser.Visible = False
txtpass.Visible = False
imgsignin.Visible = False
optagree.Visible = True
optdisagree.Visible = True
cmdcontinue.Visible = True
End Sub

Private Sub Label1_Click()

End Sub

Private Sub optagree_Click()
cmdcontinue.Enabled = True
End Sub

Private Sub optdisagree_Click()
cmdcontinue.Enabled = False
End Sub

Private Sub txtpass_Change()
imgsignin.Enabled = True
End Sub
