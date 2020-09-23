VERSION 5.00
Begin VB.Form frmQuestions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network questions"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1440
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.OptionButton optAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1440
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.OptionButton optAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Index           =   4
      Left            =   1440
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.OptionButton optAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   1440
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   5760
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Previous"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Grade Exam"
      Height          =   615
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Next"
      Height          =   615
      Index           =   0
      Left            =   7320
      TabIndex        =   8
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CheckBox chkAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   1440
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CheckBox chkAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   1440
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CheckBox chkAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CheckBox chkAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CheckBox chkAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.OptionButton optAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   1440
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Label lblQno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8880
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This original author of the code is Dave Applegate and Sachin Palewar
Option Explicit
Dim iMin As Long, iSec As Long 'Used for keeping track of remaining time.


Private Sub cmdButton_Click(Index As Integer)
 Dim iAnswer As Long
 Dim Ret As Long
 
 If oQuestions(CLng(lblQuestion.Tag)).Multiple Then
  If chkAnswer(1).Value Then iAnswer = iAnswer Or 1
  If chkAnswer(2).Value Then iAnswer = iAnswer Or 2
  If chkAnswer(4).Value Then iAnswer = iAnswer Or 4
  If chkAnswer(8).Value Then iAnswer = iAnswer Or 8
  If chkAnswer(16).Value Then iAnswer = iAnswer Or 16
 Else
  If optAnswer(1).Value Then iAnswer = 1
  If optAnswer(2).Value Then iAnswer = 2
  If optAnswer(4).Value Then iAnswer = 4
  If optAnswer(8).Value Then iAnswer = 8
  If optAnswer(16).Value Then iAnswer = 16
 End If
    
 oQuestions(CLng(lblQuestion.Tag)).UserAnswer = iAnswer
   
 Ret = GetQuestion(Index) ' Index is 0 or 1 - (cmdButton_Click(0) or cmdButton_Click(1))

    
 If Ret > 0 Then
  ShowQuestion Ret 'Ret = Question number
  If Index = 1 Then
   If oQuestions(CLng(lblQuestion.Tag)).Index = 1 Then ' Previous
    cmdButton(1).Enabled = False
   Else
    cmdButton(1).Enabled = True
   End If
    cmdButton(0).Enabled = True
    cmdButton(2).Visible = False
  Else
   If oQuestions(CLng(lblQuestion.Tag)).Index = oQuestions.Count Then   ' Next
    'Last question
    cmdButton(0).Enabled = False
    cmdButton(2).Visible = True
   Else
    cmdButton(0).Enabled = True
    cmdButton(2).Visible = False
   End If
    cmdButton(1).Enabled = True
  End If
  
 Else
  For Each oQuestion In oQuestions
   If oQuestion.Answer = oQuestion.UserAnswer Then iCorrect = iCorrect + 1
   If oQuestion.Answer <> oQuestion.UserAnswer Then iWrong = iWrong + 1
  Next
  
  iTotal = oQuestions.Count
  Unload Me
  frmResult.Show
  Set oQuestions = Nothing
 End If
End Sub

Private Sub Form_Load()
Icon = LoadPicture(App.Path & "\network.ico")
 
 Dim oDatabase As ADODB.Connection: Set oDatabase = New ADODB.Connection
 Dim oRS As ADODB.Recordset
 Dim oQuestion As mcQuestion
 Dim Ret As Long
 Dim iCount As Long
 
 Me.BackColor = vbWhite
 
 Set oQuestions = New mcQuestions

 iMin = 90
 iSec = 0
 lblTime.Caption = "90:00"
 
 oDatabase.Provider = "Microsoft.Jet.OLEDB.4.0"
 oDatabase.Open App.Path & "\TestDatabase.mdb", "admin", ""
 
 If Err Then
  oDatabase.Close: Set oDatabase = Nothing
  MsgBox "ERROR: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "There was a fatal error opening the database. Please make sure the database resides in the same directory as the application executable.", vbCritical, "Error opening database!"
  Unload Me
 End If
 Set oRS = oDatabase.Execute("SELECT * FROM Questions ORDER BY Question", , adOpenForwardOnly Or adLockReadOnly)
 
 If Not oRS.EOF Then
  Do While Not oRS.EOF
   iCount = iCount + 1
   Set oQuestion = New mcQuestion

   With oQuestion
   .Index = iCount
   .QuestionID = oRS("QID")
   .Question = "" & oRS("Question")
   .Guess1 = "" & oRS("Ans1")
   .Guess2 = "" & oRS("Ans2")
   .Guess4 = "" & oRS("Ans4")
   .Guess8 = "" & oRS("Ans8")
   .Guess16 = "" & oRS("Ans16")
   .Answer = oRS("Answer")
   .Multiple = oRS("Type")
   End With
   
   oQuestions.Add oQuestion

   Set oQuestion = Nothing
   oRS.MoveNext
  Loop
 Else
  MsgBox "No questions to ask!"
  Unload Me
 End If
  
  Ret = GetQuestion(mcNext, oQuestions.Count)
  
  
  If Ret > 0 Then ShowQuestion Ret
    Me.Show
  Exit Sub

'clear the contents of lblQuestion
lblQuestion.Caption = ""

End Sub
Private Sub Form_Unload(Cancel As Integer)
Icon = LoadPicture()
 Set oQuestions = Nothing
End Sub

Private Sub Timer1_Timer()
 'This sub Displays remaining time on the form
 If iMin = 0 And iSec = 0 Then
  MsgBox "Your score is 0.", vbOKOnly, "Time is up"
  Unload frmQuestions
  iCorrect = 0
  iWrong = 90
  frmResult.Show
  Exit Sub
 End If

 If iSec = 0 Then
  iSec = 60
  iMin = iMin - 1
 End If

 iSec = iSec - 1
 If iSec < 10 Then lblTime.Caption = Trim(Str(iMin)) + ":" + Trim("0") + Trim(Str(iSec)): Exit Sub
 lblTime.Caption = Trim(Str(iMin)) + ":" + Trim(Str(iSec))
End Sub
