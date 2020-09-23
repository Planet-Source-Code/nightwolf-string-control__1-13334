VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String Control"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Changes"
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4455
      Begin VB.OptionButton Option9 
         Caption         =   "Minimize whole sentence"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   4215
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Capitalize whole sentence"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   4215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Add a dot (.) between each character"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   4215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Minimize last letter of each word"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   4215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Minimize first letter of each word"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   4215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Capitalize last letter of each word"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   4215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Inverse each word"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   4215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Inverse whole sentence"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   4215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Capitalize first letter of each word"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Apply changes"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   2415
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         Height          =   765
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "After changes:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim NextChrSpc As Boolean
Dim TempWord As String
InitString = Text1.Text 'Remember initial string

If Option1.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 Text1.Text = Text1.Text + UCase$(Left$(TempString, 1)) 'Capitalize first letter of the string
 NextChrSpc = False
 For i% = 2 To Len(TempString) 'Go throught the whole string
  If NextChrSpc = False Then Text1.Text = Text1.Text + Mid$(TempString, i%, 1) 'Copies the letter
  If NextChrSpc = True Then Text1.Text = Text1.Text + UCase$(Mid$(TempString, i%, 1)) 'Copies the letter + Capitalizes letter
  If NextChrSpc = True Then NextChrSpc = False
  If Mid$(Text1.Text, i%, 1) = " " Then NextChrSpc = True
 Next i%
End If

If Option2.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 For i% = 1 To Len(TempString) 'Go through the whole string
  Text1.Text = Text1.Text + Mid$(TempString, Len(TempString) - (i% - 1), 1) 'Inverses string
 Next i%
End If

If Option3.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 TempWord = ""
 NextChrSpc = False
 For i% = 1 To Len(TempString) 'Go through the whole string
  If Not Mid$(TempString, i%, 1) = " " Then TempWord = TempWord + Mid$(TempString, i%, 1) 'Seeks for a word
  If Mid$(TempString, i%, 1) = " " Or i% = Len(TempString) Then
   For j% = Len(TempWord) To 1 Step -1 'Go through the whole word
    Text1.Text = Text1.Text + Mid$(TempWord, j%, 1) 'Inverse word
   Next j%
  TempWord = "" 'Clears temporary word
  End If
  If Mid$(TempString, i%, 1) = " " Then Text1.Text = Text1.Text + " " 'Add a space
 Next i%
End If

If Option4.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 NextChrSpc = False
 For i% = 1 To Len(TempString) 'Go throught the whole string
  If Mid$(TempString, i% + 1, 1) = " " Then NextChrSpc = True
  If NextChrSpc = False And Not i% = Len(TempString) Then Text1.Text = Text1.Text + Mid$(TempString, i%, 1) 'Copies the letter
  If NextChrSpc = True Or i% = Len(TempString) Then Text1.Text = Text1.Text + UCase$(Mid$(TempString, i%, 1)) 'Copies the letter + Capitalizes letter
  If NextChrSpc = True Then NextChrSpc = False
 Next i%
End If

If Option5.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 Text1.Text = Text1.Text + LCase$(Left$(TempString, 1)) 'Minimize first letter of the string
 NextChrSpc = False
 For i% = 2 To Len(TempString) 'Go throught the whole string
  If NextChrSpc = False Then Text1.Text = Text1.Text + Mid$(TempString, i%, 1) 'Copies the letter
  If NextChrSpc = True Then Text1.Text = Text1.Text + LCase$(Mid$(TempString, i%, 1)) 'Copies the letter + Minimizes letter
  If NextChrSpc = True Then NextChrSpc = False
  If Mid$(Text1.Text, i%, 1) = " " Then NextChrSpc = True
 Next i%
End If

If Option6.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 NextChrSpc = False
 For i% = 1 To Len(TempString) 'Go throught the whole string
  If Mid$(TempString, i% + 1, 1) = " " Then NextChrSpc = True
  If NextChrSpc = False And Not i% = Len(TempString) Then Text1.Text = Text1.Text + Mid$(TempString, i%, 1) 'Copies the letter
  If NextChrSpc = True Or i% = Len(TempString) Then Text1.Text = Text1.Text + UCase$(Mid$(TempString, i%, 1)) 'Copies the letter + Minimizes letter
  If NextChrSpc = True Then NextChrSpc = False
 Next i%
End If

If Option7.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 For i% = 1 To Len(TempString) 'Go through the whole string
 If Not Mid$(TempString, i%, 1) = " " Then Text1.Text = Text1.Text + Mid$(TempString, i%, 1) + "." 'Copies letter + Adds a dot (.)
 If Mid$(TempString, i%, 1) = " " Then Text1.Text = Text1.Text + " " 'Doesn't add a dot if it's a space
 Next i%
End If

If Option8.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 Text1.Text = UCase$(TempString) 'Capitalize whole string
End If

If Option9.Value = True Then
 TempString = Text1.Text
 Text1.Text = ""
 Text1.Text = LCase$(TempString) 'Minimize whole string
End If
 
Text2.Text = Text1.Text 'Show the string after changes
Text1.Text = InitString 'Displays the initial string
End Sub
