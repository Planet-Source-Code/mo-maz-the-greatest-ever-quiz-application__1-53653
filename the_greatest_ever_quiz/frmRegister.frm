VERSION 5.00
Begin VB.Form frmRegister 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register an account"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLastName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtFirstName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblLastName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblFirstName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Please enter your login details to participate in the exam:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblPasssword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Register Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The original author of this code is Rafay Mansoor
Option Explicit
Dim rs As ADODB.Recordset

Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdBack_Click()
 frmLogin.Show
 Unload Me
End Sub

Private Sub cmdSubmit_Click()
 If (txtUsername = "" Or txtPassword = "") Or (txtFirstName = "" Or txtLastName = "") Then
  MsgBox "Please complete all the fields", vbCritical = vbOKOnly, "Incomplete Login Details"
  txtUsername.SetFocus
 Else
  'add new account to login database
 With rs
  .AddNew
  !UserName = Trim(txtUsername)
  !Password = Trim(txtPassword)
  !FirstName = Trim(txtFirstName)
  !LastName = Trim(txtLastName)
  
  On Error GoTo Duplicate
  .Update
  MsgBox "Congratulations!  You can use the system now.", vbInformation, "Details entered"
    Unload Me
  frmLogin.Show
Exit Sub

Duplicate:
  If Err Then
   MsgBox txtUsername.Text & " already exists! Please try another one.", vbInformation, "Duplicate Username"
   .CancelUpdate
   txtUsername = ""
   txtPassword = ""
   txtFirstName = ""
   txtLastName = ""
   txtUsername.SetFocus
   Exit Sub
  End If
  Resume Next

 End With
 End If
End Sub

Private Sub Form_Load()
 Open_cn
'add new account to login database
Set rs = New ADODB.Recordset
rs.Open ("Select * from Login"), oCn, adOpenStatic, adLockOptimistic, _
    adCmdText
    Icon = LoadPicture(App.Path & "\network.ico")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Icon = LoadPicture()
Close_cn
End Sub
