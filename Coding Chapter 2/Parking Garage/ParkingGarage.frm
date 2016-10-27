VERSION 5.00
Begin VB.Form ParkingGarage 
   Caption         =   "Parking Garage"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCost 
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   4935
      Begin VB.PictureBox picCost 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   4635
         TabIndex        =   4
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame fraStay 
      Caption         =   "Duration of Stay"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
      Begin VB.TextBox txtTM 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblTM 
         Caption         =   "Total Minutes:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraTime 
      Caption         =   "Arrival Time"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4935
      Begin VB.OptionButton optA6 
         Caption         =   "After 6AM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton opt826 
         Caption         =   "8AM - 6AM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optB8 
         Caption         =   "Before 8AM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraDay 
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.OptionButton optWE 
         Caption         =   "Weekend"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optWD 
         Caption         =   "Week Day"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ParkingGarage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: October 3 2013
'Purpose: Play around with more variables
Option Explicit

Private Sub Form_Load()
optWD.Value = True
txtTM.Enabled = False
lblTM.Enabled = False
fraStay.Enabled = False
End Sub

Private Sub optA6_Click()
txtTM.Enabled = False
lblTM.Enabled = False
fraStay.Enabled = False
picCost.Cls
picCost.Print "$6"
End Sub

Private Sub optB8_Click()
txtTM.Enabled = False
lblTM.Enabled = False
fraStay.Enabled = False
picCost.Cls
picCost.Print "$15"
End Sub

Private Sub optWD_Click()
optB8.Enabled = True
opt826.Enabled = True
optA6.Enabled = True
txtTM.Enabled = True
lblTM.Enabled = True
fraTime.Enabled = True
fraStay.Enabled = True
picCost.Cls
End Sub

Private Sub optWE_Click()
optB8.Enabled = False
opt826.Enabled = False
optA6.Enabled = False
fraTime.Enabled = False

txtTM.Enabled = False
lblTM.Enabled = False
fraStay.Enabled = False
picCost.Cls
picCost.Print "$5"
End Sub

Private Sub opt826_Click()
txtTM.Enabled = True
lblTM.Enabled = True
fraStay.Enabled = True
picCost.Cls
End Sub

Private Sub txtTM_Change()
'declare
Dim sglTM As Single
Dim sglCost As Single
'intialize
sglTM = 0
sglCost = 0
'input
sglTM = Val(txtTM.Text)
'process/cal
sglCost = (sglTM \ 30) * 2.75
If sglTM Mod 30 > 0 Then sglCost = (sglTM \ 30 + 1) * 2.75

'output
picCost.Cls
picCost.Print Format(sglCost, "$0.00")
picCost.Print sglCost
End Sub
