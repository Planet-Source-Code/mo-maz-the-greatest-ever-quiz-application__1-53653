VERSION 5.00
Begin VB.Form frmInstruction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructions"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Test"
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lblGreetings 
      Caption         =   "Greetings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblInstructions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Testing Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   2145
   End
End
Attribute VB_Name = "frmInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was modified using one of the Deitel exercises
Option Explicit                          'General declaration
Dim mFileSysObj As New FileSystemObject  'General declaration
Dim mFile As File                        'General declaration
Dim mTxtStream As TextStream             'General declaration
Dim instructions As String               'General declaration
Private Sub cmdStart_Click()
 frmQuestions.Show
 Unload Me
End Sub

Private Sub Form_Load()
Icon = LoadPicture(App.Path & "\network.ico")
 lblGreetings.Caption = ""
 
 On Error GoTo MissingFile
 'Get the file
 Set mFile = mFileSysObj.GetFile(App.Path + "\Instructions.txt")

 'Open a text stream for reading to the file
 Set mTxtStream = mFile.OpenAsTextStream(ForReading)
   
 'Read the data
 instructions = mTxtStream.ReadAll
 
 'Get the first name of the candidate
 n = InStr(strName, " ")
 strFirst = Left(strName, n - 1)
 
 'Greets candidate by First Name
 lblGreetings.Caption = "Welcome " & strFirst & "!"
                                                        
 'Place only the String portion representing the
 'name in the TextBox.
 lblInstructions.Caption = instructions
 Exit Sub
 
MissingFile:
MsgBox "Please make sure the instructions.txt resides in the same directory as the executable file." & vbCrLf & vbCrLf & "This program will close down.", vbExclamation, "Instructions file missing"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Icon = LoadPicture()
End Sub
