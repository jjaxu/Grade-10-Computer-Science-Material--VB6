VERSION 5.00
Begin VB.Form frmVowelCounter 
   Caption         =   "Vowel Counter"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkY 
      Caption         =   "Count letter ""Y"" as vowel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   7695
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4695
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   7695
   End
   Begin VB.Label lblTitle 
      Caption         =   "Vowel Counter"
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
      Width           =   7575
   End
End
Attribute VB_Name = "frmVowelCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: November 26 2013
'Purpose: To use Decisions with loop with strings
Option Explicit

Private Sub chkY_Click()
lstOutput.Clear
End Sub

Private Sub cmdCount_Click()
'declare
Dim intVA As Integer
Dim intVE As Integer
Dim intVI As Integer
Dim intVO As Integer
Dim intVU As Integer
Dim intVY As Integer

Dim intVTotal As Integer
Dim intAllLetters As Integer
Dim intConsonants As Integer


Dim strInput As String
Dim intCounter As Integer
'initialize
intVA = 0
intVE = 0
intVI = 0
intVO = 0
intCounter = 0
intVTotal = 0

intVU = 0
intVY = 0
intAllLetters = 0
strInput = ""

'input
strInput = LCase(txtInput.Text)
'process / calculation
lstOutput.Clear
If chkY.Value = 1 Then
'Vowel A
Do While InStr(strInput, "a") <> 0
    intCounter = InStr(strInput, "a")
    strInput = Mid(strInput, intCounter + 1)
    intVA = intVA + 1
Loop
lstOutput.AddItem "A - " & intVA

'Vowel E
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "e") <> 0
    intCounter = InStr(strInput, "e")
    strInput = Mid(strInput, intCounter + 1)
    intVE = intVE + 1
Loop
lstOutput.AddItem "E - " & intVE
    
'Vowel I
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "i") <> 0
    intCounter = InStr(strInput, "i")
    strInput = Mid(strInput, intCounter + 1)
    intVI = intVI + 1
Loop
lstOutput.AddItem "I - " & intVI

'Vowel O
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "o") <> 0
    intCounter = InStr(strInput, "o")
    strInput = Mid(strInput, intCounter + 1)
    intVO = intVO + 1
Loop
lstOutput.AddItem "O - " & intVO

'Vowel U
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "u") <> 0
    intCounter = InStr(strInput, "u")
    strInput = Mid(strInput, intCounter + 1)
    intVU = intVU + 1
Loop
lstOutput.AddItem "U - " & intVU

'Vowel Y
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "y") <> 0
    intCounter = InStr(strInput, "y")
    strInput = Mid(strInput, intCounter + 1)
    intVY = intVY + 1
Loop
lstOutput.AddItem "Y - " & intVY
intVTotal = intVA + intVE + intVI + intVO + intVU + intVY
lstOutput.AddItem "Total vowels: " & intVTotal

Else
'Vowel A
Do While InStr(strInput, "a") <> 0
    intCounter = InStr(strInput, "a")
    strInput = Mid(strInput, intCounter + 1)
    intVA = intVA + 1
Loop
lstOutput.AddItem "A - " & intVA

'Vowel E
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "e") <> 0
    intCounter = InStr(strInput, "e")
    strInput = Mid(strInput, intCounter + 1)
    intVE = intVE + 1
Loop
lstOutput.AddItem "E - " & intVE
    
'Vowel I
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "i") <> 0
    intCounter = InStr(strInput, "i")
    strInput = Mid(strInput, intCounter + 1)
    intVI = intVI + 1
Loop
lstOutput.AddItem "I - " & intVI

'Vowel O
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "o") <> 0
    intCounter = InStr(strInput, "o")
    strInput = Mid(strInput, intCounter + 1)
    intVO = intVO + 1
Loop
lstOutput.AddItem "O - " & intVO

'Vowel U
strInput = LCase(txtInput.Text)
intCounter = 0
Do While InStr(strInput, "u") <> 0
    intCounter = InStr(strInput, "u")
    strInput = Mid(strInput, intCounter + 1)
    intVU = intVU + 1
Loop
lstOutput.AddItem "U - " & intVU
intVTotal = intVA + intVE + intVI + intVO + intVU
lstOutput.AddItem "Total vowels: " & intVTotal
End If

'Consonants
strInput = LCase(txtInput.Text)
intCounter = 1
Do While Len(strInput) > 0
    If Left(strInput, 1) >= "a" And Left(strInput, 1) <= "z" Then
        intAllLetters = intAllLetters + 1
    End If
    strInput = Mid(strInput, intCounter + 1)
Loop
intConsonants = intAllLetters - intVTotal

lstOutput.AddItem "Total Consonants: " & intConsonants
End Sub

Private Sub txtInput_Change()
lstOutput.Clear
End Sub
