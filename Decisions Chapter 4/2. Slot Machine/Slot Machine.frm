VERSION 5.00
Begin VB.Form SlotMachine 
   Caption         =   "Slot Machine"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay the debt"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   6240
      Width           =   8295
   End
   Begin VB.CommandButton cmdBorrow 
      Caption         =   "Borrow money from the bank"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   8295
   End
   Begin VB.CommandButton cmdPull 
      Caption         =   "PULL!!!! (Cost $1 per pull)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Image imgSlot1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image imgSlot2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image imgSlot3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   7
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblOwe 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   8295
   End
   Begin VB.Label lblMoney 
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
      Height          =   975
      Left            =   5640
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblWallet 
      Caption         =   "Your Wallet"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblSlotM 
      Alignment       =   2  'Center
      Caption         =   "Slot Machine"
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
      Width           =   5055
   End
End
Attribute VB_Name = "SlotMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: November 4th 2013
'Purpose: To learn more about decisions
'global variable
Dim intMoney As Integer
Dim intOwe As Integer
Option Explicit

Private Sub cmdBorrow_Click()
intMoney = intMoney + 20
intOwe = intOwe + 20
cmdPull.Enabled = True
lblMoney.Caption = "$" & intMoney
lblOwe.Caption = "You owe the bank $" & intOwe
cmdBorrow.Enabled = False
lblStatus.Caption = "Press 'PULL' to play!"

End Sub

Private Sub cmdPay_Click()
intMoney = intMoney - 20
intOwe = intOwe - 20
lblMoney.Caption = "$" & intMoney
lblOwe.Caption = "You owe the bank $" & intOwe
    
If intMoney > 20 And intOwe >= 20 Then
    cmdPay.Enabled = True
Else: cmdPay.Enabled = False
End If
End Sub

Private Sub cmdPull_Click()
'declare
Dim intNumber1 As Integer
Dim intNumber2 As Integer
Dim intNumber3 As Integer
'initialize
intNumber1 = 0
intNumber2 = 0
intNumber3 = 0
'process/calculation
intNumber1 = Int(3 * Rnd) + 1
intNumber2 = Int(3 * Rnd) + 1
intNumber3 = Int(3 * Rnd) + 1
intMoney = intMoney - 1

If intNumber1 = 1 And intNumber2 = 1 And intNumber3 = 1 Then
    intMoney = intMoney + 4
    lblStatus = "Small Win! You win $4!"
ElseIf intNumber1 = 2 And intNumber2 = 2 And intNumber3 = 2 Then
    intMoney = intMoney + 8
    lblStatus = "Gold Rush! You win $8!"
ElseIf intNumber1 = 3 And intNumber2 = 3 And intNumber3 = 3 Then
    intMoney = intMoney + 12
    lblStatus = "JACKPOT! You win $12!"
Else: lblStatus = "Try again"
End If
'output
If intNumber1 = 1 Then
    imgSlot1.Picture = LoadPicture(App.Path & "\Silver Coin.bmp")

ElseIf intNumber1 = 2 Then
    imgSlot1.Picture = LoadPicture(App.Path & "\Gold Coin.bmp")
   
Else: imgSlot1.Picture = LoadPicture(App.Path & "\Diamond.bmp")
End If


If intNumber2 = 1 Then
    imgSlot2.Picture = LoadPicture(App.Path & "\Silver Coin.bmp")

ElseIf intNumber2 = 2 Then
    imgSlot2.Picture = LoadPicture(App.Path & "\Gold Coin.bmp")
   
Else: imgSlot2.Picture = LoadPicture(App.Path & "\Diamond.bmp")
End If

If intNumber3 = 1 Then
    imgSlot3.Picture = LoadPicture(App.Path & "\Silver Coin.bmp")

ElseIf intNumber3 = 2 Then
    imgSlot3.Picture = LoadPicture(App.Path & "\Gold Coin.bmp")
   
Else: imgSlot3.Picture = LoadPicture(App.Path & "\Diamond.bmp")
End If

If intMoney <= 0 Then
    lblStatus = "Insufficient funds!!!"
End If

lblMoney.Caption = "$" & intMoney

If intMoney <= 0 Then
    cmdPull.Enabled = False
    cmdBorrow.Enabled = True
Else: cmdBorrow.Enabled = False
End If

If intMoney > 20 And intOwe >= 20 Then
    cmdPay.Enabled = True
Else: cmdPay.Enabled = False
End If

End Sub

Private Sub Form_Load()
Randomize
lblStatus.Caption = "Press 'PULL' to play!"
intMoney = 20
intOwe = 0
lblMoney.Caption = "$" & intMoney
cmdBorrow.Enabled = False
cmdPay.Enabled = False

End Sub
