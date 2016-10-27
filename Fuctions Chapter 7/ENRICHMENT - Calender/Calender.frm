VERSION 5.00
Begin VB.Form frmCalender 
   Caption         =   "Calender"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCalender 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      ScaleHeight     =   2955
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   5655
   End
   Begin VB.ComboBox cboDay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Calender.frx":0000
      Left            =   4800
      List            =   "Calender.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Calender.frx":0004
      Left            =   3480
      List            =   "Calender.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Caption         =   "Day:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      Caption         =   "Month:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblDate 
      Caption         =   "Enter date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Caption         =   "Calender"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jackie Xu
'Date: December 19 2013
'Purpose: Enrichment question, more functions
Option Explicit
'GV
Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer

Private Sub cboDay_Click()
picCalender.Cls
'input
intYear = Trim(Val(txtYear.Text))
intMonth = Trim(Val(cboMonth.Text))
intDay = Trim(Val(cboDay.Text))

If intYear / 2 <> 0 And intYear > 0 And cboMonth.Text <> "" And cboDay.Text <> "" Then
    cmdDisplay.Enabled = True
Else
    cmdDisplay.Enabled = False
End If
End Sub

Private Sub cboMonth_click()
'input
intYear = Trim(Val(txtYear.Text))
intMonth = Trim(Val(cboMonth.Text))
intDay = Trim(Val(cboDay.Text))

cboDay.Clear
picCalender.Cls
Dim intCount As Integer

For intCount = 1 To 28
    cboDay.AddItem intCount
Next intCount
If intMonth = 2 And leapYear = True Then
    cboDay.AddItem "29"
ElseIf (intMonth Mod 2 = 1 And intMonth <= 7) Or _
(intMonth Mod 2 = 0 And intMonth >= 8) Then
    cboDay.AddItem "29"
    cboDay.AddItem "30"
    cboDay.AddItem "31"
ElseIf intMonth <> 2 Then
    cboDay.AddItem "29"
    cboDay.AddItem "30"
End If

If intYear / 2 <> 0 And intYear > 0 And cboMonth.Text <> "" And cboDay.Text <> "" Then
    cmdDisplay.Enabled = True
Else
    cmdDisplay.Enabled = False
End If
End Sub
Public Sub cmdDisplay_Click()

Dim intDayCount As Integer
Dim intCounter As Integer
Dim intCounter2 As Integer
Dim strSpace As String


Dim strCurrent As String
Dim strResult As String

intCounter = 0
intDayCount = 0
intCounter2 = 0
strSpace = ""
strCurrent = ""
strResult = ""

intYear = Trim(Val(txtYear.Text))
intMonth = Trim(Val(cboMonth.Text))
intDay = Trim(Val(cboDay.Text))

picCalender.Cls
picCalender.Print dayName & ", " & intDay & " " & monthName & " " & intYear
picCalender.Print ""
picCalender.Print "Sun  Mon  Tue  Wed  Thu  Fri  Sat"


'Do While intDayCount < maxDays
'1st row
    For intCounter = zeller(intMonth, 1, intYear) + 1 To 7
         intDayCount = intDayCount + 1
         strResult = strResult & "  " & intDayCount & "  "
    Next intCounter
    For intCounter2 = 1 To zeller(intMonth, 1, intYear)
        strSpace = strSpace & "     "
    Next intCounter2
        picCalender.Print strSpace & strResult

'2nd row
    strResult = ""
    For intCounter = 1 To 7
        intDayCount = intDayCount + 1
            If intDayCount < 10 Then
                strResult = strResult & "  " & intDayCount & "  "
            Else
                strResult = strResult & " " & intDayCount & "  "
            End If
    Next intCounter
        picCalender.Print strResult
'3rd row
    strResult = ""
    For intCounter = 1 To 7
         intDayCount = intDayCount + 1
            If intDayCount < 10 Then
                strResult = strResult & "  " & intDayCount & "  "
            Else
                strResult = strResult & " " & intDayCount & "  "
            End If
    Next intCounter
        picCalender.Print strResult
'4th row
    strResult = ""
    For intCounter = 1 To 7
         intDayCount = intDayCount + 1
            If intDayCount < 10 Then
                strResult = strResult & "  " & intDayCount & "  "
            Else
                strResult = strResult & " " & intDayCount & "  "
            End If
    Next intCounter
        picCalender.Print strResult
If zeller(intMonth, 1, intYear) <= 4 Then
'5th row (without 6th)
    strResult = ""
    For intCounter = intDayCount + 1 To maxDays
        intDayCount = intDayCount + 1
        strResult = strResult & " " & intDayCount & "  "
    Next intCounter
        picCalender.Print strResult
ElseIf zeller(intMonth, 1, intYear) > 4 And intMonth = 2 Then
'Special Feb (5 rows even zeller > 4)
    strResult = ""
    For intCounter = intDayCount + 1 To maxDays
        intDayCount = intDayCount + 1
            If intDayCount < 10 Then
                strResult = strResult & "  " & intDayCount & "  "
            Else
                strResult = strResult & " " & intDayCount & "  "
            End If
    Next intCounter
        picCalender.Print strResult
Else
'5th row (with 6th)
    strResult = ""
    For intCounter = 1 To 7
        intDayCount = intDayCount + 1
            If intDayCount < 10 Then
                strResult = strResult & "  " & intDayCount & "  "
            Else
                strResult = strResult & " " & intDayCount & "  "
            End If
    Next intCounter
        picCalender.Print strResult
    
'6th row
        strResult = ""
    For intCounter = intDayCount + 1 To maxDays
        intDayCount = intDayCount + 1
        strResult = strResult & " " & intDayCount & "  "
    Next intCounter
         picCalender.Print strResult
End If

'Loop

End Sub

Public Function leapYear() As Boolean
If intYear Mod 4 = 0 Then
    leapYear = True
Else
    leapYear = False
End If
End Function

Public Function maxDays() As Integer
Dim intTotalDays As Integer
If intMonth = 2 Then
    If leapYear = True Then
        intTotalDays = 29
    Else
        intTotalDays = 28
    End If
ElseIf (intMonth Mod 2 = 1 And intMonth <= 7) Or _
(intMonth Mod 2 = 0 And intMonth >= 8) Then
    intTotalDays = 31
Else
    intTotalDays = 30
End If
maxDays = intTotalDays
End Function

Public Function monthName() As String
If intMonth = 1 Then
    monthName = "January"
ElseIf intMonth = 2 Then
    monthName = "February"
ElseIf intMonth = 3 Then
    monthName = "March"
ElseIf intMonth = 4 Then
    monthName = "April"
ElseIf intMonth = 5 Then
    monthName = "May"
ElseIf intMonth = 6 Then
    monthName = "June"
ElseIf intMonth = 7 Then
    monthName = "July"
ElseIf intMonth = 8 Then
    monthName = "August"
ElseIf intMonth = 9 Then
    monthName = "September"
ElseIf intMonth = 10 Then
    monthName = "October"
ElseIf intMonth = 11 Then
    monthName = "November"
ElseIf intMonth = 12 Then
    monthName = "December"
End If
End Function

Private Sub Form_Load()
cboDay.Enabled = False
cboMonth.Enabled = False
cmdDisplay.Enabled = False
End Sub

Private Sub txtYear_Change()
intYear = Trim(Val(txtYear.Text))
If intYear / 2 = 0 Or intYear < 0 Then
    cboMonth.Enabled = False
    cboDay.Enabled = False
Else
    cboMonth.Enabled = True
    cboDay.Enabled = True
End If


cboDay.Clear
picCalender.Cls
Dim intCount As Integer
For intCount = 1 To 28
    cboDay.AddItem intCount
Next intCount
If intMonth = 2 And leapYear = True Then
    cboDay.AddItem "29"
ElseIf (intMonth Mod 2 = 1 And intMonth <= 7) Or _
(intMonth Mod 2 = 0 And intMonth >= 8) Then
    cboDay.AddItem "29"
    cboDay.AddItem "30"
    cboDay.AddItem "31"
ElseIf intMonth <> 2 Then
    cboDay.AddItem "29"
    cboDay.AddItem "30"
End If

If intYear / 2 <> 0 And intYear > 0 And cboMonth.Text <> "" And cboDay.Text <> "" Then
    cmdDisplay.Enabled = True
Else
    cmdDisplay.Enabled = False
End If
End Sub

Public Function dayName() As String
If zeller(intMonth, intDay, intYear) = 0 Then
    dayName = "Sunday"
ElseIf zeller(intMonth, intDay, intYear) = 1 Then
    dayName = "Monday"
ElseIf zeller(intMonth, intDay, intYear) = 2 Then
    dayName = "Tuesday"
ElseIf zeller(intMonth, intDay, intYear) = 3 Then
    dayName = "Wednesday"
ElseIf zeller(intMonth, intDay, intYear) = 4 Then
    dayName = "Thursday"
ElseIf zeller(intMonth, intDay, intYear) = 5 Then
    dayName = "Friday"
ElseIf zeller(intMonth, intDay, intYear) = 6 Then
    dayName = "Saturday"
End If
End Function

Public Function zeller(month, day, year) As Integer
Dim m As Integer
Dim y As Integer
Dim p As Integer
Dim r As Integer

intYear = Trim(Val(txtYear.Text))
intMonth = Trim(Val(cboMonth.Text))
intDay = Trim(Val(cboDay.Text))

m = month - 2
y = year
If m < 0 Then
    m = m + 12
    y = y - 1
End If
p = y \ 100
r = y Mod 100

zeller = (day + (26 * m - 2) \ 10 + r + r \ 4 + p \ 4 + 5 * p) Mod 7
End Function

Public Function firstDay() As Integer
firstDay = zeller(intMonth, 1, intYear)
End Function
