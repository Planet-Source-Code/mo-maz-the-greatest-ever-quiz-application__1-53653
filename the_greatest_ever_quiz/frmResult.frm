VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmResult 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Test Result"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -465
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Score Report"
      Height          =   495
      Left            =   2520
      TabIndex        =   21
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Click here to exit"
      Height          =   495
      Left            =   7800
      TabIndex        =   18
      Top             =   8160
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar PBar2 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin VB.Label lblPassFail 
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
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Passed/Failed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Line Line5 
      X1              =   2520
      X2              =   9240
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   17
      Top             =   1800
      Width           =   75
   End
   Begin VB.Shape Shape5 
      Height          =   495
      Left            =   840
      Top             =   1680
      Width           =   10335
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Date appeared :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6000
      TabIndex        =   16
      Top             =   1200
      Width           =   1710
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   5880
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   840
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   15
      Top             =   6480
      Width           =   75
   End
   Begin VB.Label lblPercent 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   14
      Top             =   5880
      Width           =   75
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Total Test marks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   13
      Top             =   6480
      Width           =   1770
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The total marks obtained by candidate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   12
      Top             =   5880
      Width           =   4005
   End
   Begin VB.Line Line4 
      X1              =   2520
      X2              =   9240
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   9240
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   7560
      Y1              =   4440
      Y2              =   7560
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   9240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblWrong 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   11
      Top             =   5280
      Width           =   75
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   10
      Top             =   4680
      Width           =   75
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   9
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The number of Wrong Answer is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   8
      Top             =   5220
      Width           =   75
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The number of Right Answer is "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   7
      Top             =   4620
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   6
      Top             =   3120
      Width           =   330
   End
   Begin VB.Label lblMax 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "60%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6720
      TabIndex        =   5
      Top             =   2400
      Width           =   450
   End
   Begin VB.Label LabelTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Score Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4695
      TabIndex        =   4
      Top             =   360
      Width           =   1860
   End
   Begin VB.Label lblOMarks 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Marks Obtained"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   1650
   End
   Begin VB.Label lblRMarks 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Marks Required"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   840
      Top             =   2280
      Width           =   10335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   2520
      Top             =   4440
      Width           =   6735
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The original author of this code is Catejan D'Souza
Option Explicit
Dim oRS As ADODB.Recordset
Dim dDate As Date
Dim iPercent As Long
Dim strStatus As String
Private Sub cmdExit_Click()
 n = InStr(strName, " ")
 strFirst = Left(strName, n - 1)
 'Display the message when the candidate exits
 MsgBox strFirst & "," & " you are closing the application. " & vbCrLf & "If you need to take another exam, you need log in again.", , "Closing down"
 Close_cn
 End
End Sub

Private Sub cmdPrint_Click()
 On Error GoTo Handler
  'Print the score report
  PrintForm
  'If there is no printer installed show the error message
Handler:
If Err Then
 MsgBox Err.Description, vbCritical, "Printer"
 End If
End Sub

Private Sub Form_Load()
 'Displays a mix of the date and time
 'this data makes the Score table unique in the database
 dDate = Now()

 'This select statement is used to display the status
 'of a candidate's test, either Passed or Failed.
 'The student must answer 36 questions correctly to pass
 Select Case iCorrect
  Case Is > 36
   strStatus = "Passed"
   lblPassFail.Caption = strStatus
  Case Else
   strStatus = "Failed"
  lblPassFail.Caption = strStatus
 End Select

 'This label will present the full name of the candidate
 lblName.Caption = "Candidate's full name:  " + strName
 'Shows the number of correct answers
 lblRight.Caption = iCorrect
 'Shows the number of incorrect answers
 lblWrong.Caption = iWrong

 'If the score is greater than 0, then calculate the percentage score
 If iCorrect > 0 Then
  iPercent = iCorrect / iTotal * 100
 Else
  'If the score is 0, then the perecentage is 0
  iPercent = 0
 End If

 'Show the first Progress Bar the value of
 '60%, indicating pass score
 PBar1.Value = 60
 'Indicate the candidate's percentage score
 PBar2.Value = iPercent
 'Convert the percentage score into a string to display
 'the percentage in a label
 Label5.Caption = Str(iPercent) + "%"
 'Display the score in a progress bar
 Label5.Move (3240 + iPercent * 68.4)
 'Convert the percentage score into a string to display
 'the percentage in a label in the table
 lblPercent.Caption = Str(iPercent) & "%"
 'Display the total number of questions available
 lblTotal.Caption = iTotal
 'Display the date in a string format in a label
 lblDate.Caption = "Date appeared : " + CStr(dDate)

 'Open a database connection to store the values
 'in the Score table
 Open_cn

 'Set the recordset object to the Score table
 Set oRS = New ADODB.Recordset
 oRS.Open ("SELECT * FROM Score"), oCn, adOpenStatic, adLockOptimistic, _
   adCmdText
   
 'Store all the relevant data to the Score table
 With oRS
  .AddNew
  !UserName = strUserName
  !TestDate = dDate
  !CorrectAns = iCorrect
  !WrongAns = iWrong
  !Score = iPercent
  !Status = strStatus
 .Update
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Icon = LoadPicture()
End Sub
