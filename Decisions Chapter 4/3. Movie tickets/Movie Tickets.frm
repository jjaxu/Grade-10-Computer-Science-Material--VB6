VERSION 5.00
Begin VB.Form frmTickets 
   Caption         =   "Ticket  Sales"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4320
      TabIndex        =   9
      Top             =   5400
      Width           =   3975
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Width           =   3975
   End
   Begin VB.ComboBox cboTickets 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4320
      TabIndex        =   7
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label lblCostTitle 
      Caption         =   "Total Cost:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   6
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label lblTicketNumber 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lblTicketsTitle 
      Caption         =   "# of Tickets:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblChoose 
      Caption         =   "Choose your ticket:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblTitle 
      Caption         =   "Ticket Sales"
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
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date" November 10 2013
'Purpose: To learn how to use combo boxes with decisions
Option Explicit
'Declare (GLOBAL)
Dim intTickets As Integer
Dim sglCost As Single

Private Sub cmdClear_Click()
lstOutput.Clear
intTickets = 0
sglCost = 0
lblTicketNumber.Caption = intTickets
lblTotalCost.Caption = "$" & sglCost
End Sub

Private Sub Form_Load()
cboTickets.AddItem "Adult ($11.00)"
cboTickets.AddItem "Senior ($8.00)"
cboTickets.AddItem "Student ($6.50)"
cboTickets.AddItem "Child ($4.75)"
cboTickets.AddItem "Coupon (Free)"
cboTickets.Text = "Adult ($11.00)"
'initialize
intTickets = 0
sglCost = 0
lblTicketNumber.Caption = "0"
lblTotalCost.Caption = "$0"
End Sub

Private Sub cmdAdd_Click()
'Process
intTickets = intTickets + 1
If cboTickets.Text = "Adult ($11.00)" Then
    sglCost = sglCost + 11
    lstOutput.AddItem "Adult"
ElseIf cboTickets.Text = "Senior ($8.00)" Then
    sglCost = sglCost + 8
    lstOutput.AddItem "Senior"
ElseIf cboTickets.Text = "Student ($6.50)" Then
    sglCost = sglCost + 6.5
    lstOutput.AddItem "Student"
ElseIf cboTickets.Text = "Child ($4.75)" Then
    sglCost = sglCost + 4.75
    lstOutput.AddItem "Child"
Else: lstOutput.AddItem "Coupon"
End If
'output
lblTicketNumber.Caption = intTickets
lblTotalCost.Caption = "$" & sglCost
End Sub

