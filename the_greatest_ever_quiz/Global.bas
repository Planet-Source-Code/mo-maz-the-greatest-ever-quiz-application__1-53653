Attribute VB_Name = "global"
Option Explicit
 Public oQuestion As mcQuestion
 Public iCorrect As Long
 Public iWrong As Long
 Public iTotal As Long
 Public strUserName As String
 Public strName As String
 Public oCn As ADODB.Connection
 Public strFirst As String
 Public n As Long
Public Sub Open_cn()
 Set oCn = New ADODB.Connection
 oCn.CursorLocation = adUseClient
 oCn.Provider = "Microsoft.Jet.OLEDB.4.0"
 oCn.Properties("Data Source") = App.Path & "\TestDatabase.mdb"
 oCn.Open
End Sub
Public Sub Close_cn()
 oCn.Close
 Set oCn = Nothing
End Sub

