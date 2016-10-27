VERSION 5.00
Begin VB.Form Receipt 
   Caption         =   "Receipt"
   ClientHeight    =   12195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   12195
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picQT 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   27
      Top             =   8160
      Width           =   2295
   End
   Begin VB.PictureBox picFinal 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   26
      Top             =   9480
      Width           =   2295
   End
   Begin VB.PictureBox picHST 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   25
      Top             =   8160
      Width           =   2295
   End
   Begin VB.PictureBox picST 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   24
      Top             =   6960
      Width           =   2295
   End
   Begin VB.PictureBox picT4 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   23
      Top             =   5760
      Width           =   2295
   End
   Begin VB.PictureBox picT3 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   22
      Top             =   4440
      Width           =   2295
   End
   Begin VB.PictureBox picT2 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   21
      Top             =   3120
      Width           =   2295
   End
   Begin VB.PictureBox picT1 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2235
      TabIndex        =   20
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtUC4 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   16
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtUC3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   15
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtUC2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   14
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtUC1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   13
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtQ4 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   12
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtQ3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   11
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtQ2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtQ1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Calculate Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   4
      Top             =   10800
      Width           =   8895
   End
   Begin VB.Label lblItemTotal 
      Caption         =   "Total items:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   28
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label lblFinal 
      Caption         =   "Final:"
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
      Left            =   4800
      TabIndex        =   19
      Top             =   9360
      Width           =   2295
   End
   Begin VB.Label lblHST 
      Caption         =   "HST:"
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
      Left            =   4800
      TabIndex        =   18
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Label lblStotal 
      Caption         =   "Sub Total:"
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
      Left            =   4800
      TabIndex        =   17
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label lblItem4 
      Caption         =   "4:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   8
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lblItem3 
      Caption         =   "3:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblItem2 
      Caption         =   "2:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblItem1 
      Caption         =   "1:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblUnitCost 
      Caption         =   "Unit Cost ($)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblQuantity 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblItem 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: October 1st 2013
'Purpose: Play around w/ variables
Option Explicit

Private Sub cmdTotal_Click()
'Declare
Dim SglQ1 As Single
Dim SglUC1 As Single
Dim SglT1 As Single

Dim SglQ2 As Single
Dim SglUC2 As Single
Dim SglT2 As Single

Dim SglQ3 As Single
Dim SglUC3 As Single
Dim SglT3 As Single

Dim SglQ4 As Single
Dim SglUC4 As Single
Dim SglT4 As Single

Dim SglQT As Single
Dim SglST As Single
Dim SglHST As Single
Dim SglFinal As Single
'Intialize
SglQ1 = 0
SglUC1 = 0
SglQ2 = 0
SglUC2 = 0
SglQ3 = 0
SglUC3 = 0
SglQ4 = 0
SglUC4 = 0
SglQT = 0
SglST = 0
SglHST = 0
SglFinal = 0
'input
SglQ1 = Val(txtQ1.Text)
SglUC1 = Val(txtUC1.Text)
SglQ2 = Val(txtQ2.Text)
SglUC2 = Val(txtUC2.Text)
SglQ3 = Val(txtQ3.Text)
SglUC3 = Val(txtUC3.Text)
SglQ4 = Val(txtQ4.Text)
SglUC4 = Val(txtUC4.Text)
'process
SglQT = SglQ1 + SglQ2 + SglQ3 + SglQ4

SglT1 = SglQ1 * SglUC1
SglT2 = SglQ2 * SglUC2
SglT3 = SglQ3 * SglUC3
SglT4 = SglQ4 * SglUC4

SglST = SglT1 + SglT2 + SglT3 + SglT4
SglHST = SglST * 0.13
SglFinal = SglST + SglHST
'output
picQT.Cls
picT1.Cls
picT2.Cls
picT3.Cls
picT4.Cls
picST.Cls
picHST.Cls
picFinal.Cls

picQT.Print SglQT

picT1.Print Format(SglT1, "  $0.00")
picT1.Print SglT1

picT2.Print Format(SglT2, "  $0.00")
picT2.Print SglT2

picT3.Print Format(SglT3, "  $0.00")
picT3.Print SglT3

picT4.Print Format(SglT4, "  $0.00")
picT4.Print SglT4

picST.Print Format(SglST, "  $0.00")
picST.Print SglST

picHST.Print Format(SglHST, "  $0.00")
picHST.Print SglHST

picFinal.Print Format(SglFinal, "  $0.00")
picFinal.Print SglFinal

picHST.Print SglHST
End Sub

