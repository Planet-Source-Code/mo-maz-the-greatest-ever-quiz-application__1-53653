VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
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
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblPasssword 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   1920
      TabIndex        =   3
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
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please login:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
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
      Caption         =   "Testing Engine"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The original author of this code is John Overton.
'I had changed some parts of the code
Option Explicit
Dim oRS As ADODB.Recordset

Private Sub cmdExit_Click()
 End
End Sub

Private Sub cmdLogin_Click()
On Error GoTo ErrHandler
 'check for blank fields
 If txtUsername.Text = "" Or IsNull(txtUsername.Text) = True Then
  Call MsgBox("Username must be entered.", vbOKOnly, "Username")
  txtUsername.SetFocus
 Exit Sub
 End If
 
 If txtPassword.Text = "" Or IsNull(txtPassword.Text) = True Then
  Call MsgBox("Password must be entered.", vbOKOnly, "Password")
  txtPassword.SetFocus
 Exit Sub
 End If
    
 'Connects to the MS Access database
 Open_cn
 
 'Reference the recordset object to the Login table
 Set oRS = New ADODB.Recordset
 oRS.Open ("Select * from Login Where Username= '" & txtUsername.Text & "'"), oCn, adOpenStatic, adLockOptimistic, _
  adCmdText
 
 'If the password is incorrect, display the associated error
 If txtPassword.Text <> oRS.Fields("Password") Then
  Call MsgBox("Incorrect Password", vbOKOnly, "Login Error")
  txtPassword.Text = ""
  txtPassword.SetFocus
 Exit Sub
 Else
  'Display the frmInstruction form
  strUserName = txtUsername.Text 'May need in the future project
  strName = oRS.Fields("FirstName") & " " & oRS.Fields("LastName")
  frmInstruction.Show
  Unload Me
 End If
  'Close the database connection
  Close_cn
 Exit Sub

ErrHandler:
 'Display the incorrect username error if an incorrect username is entered
 Call MsgBox("Incorrect Username", vbOKOnly, "Login Error")
 'Clear both fields
 txtUsername.Text = ""
 txtPassword.Text = ""
 'Set the cursor to the Username field
 txtUsername.SetFocus
 Exit Sub
End Sub
Private Sub cmdRegister_Click()
 'Display the frmRegister form if it's a new candidate
 frmRegister.Show
 'Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo NoIcon
Icon = LoadPicture(App.Path & "\network.ico")
Exit Sub

NoIcon:
MsgBox "Please make sure the network.ico exists in the same directory as the executable file." & vbCrLf & vbCrLf & "This program will close down.", vbExclamation, "Icon file missing"
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Icon = LoadPicture()
End Sub
