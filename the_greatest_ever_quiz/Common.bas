Attribute VB_Name = "Common"
Option Explicit

 Public oQuestions As mcQuestions

 #If False Then
 Public mcNext: Public mcPrevious
 #End If
 Public Enum mcGetQuestionTypes
  mcNext = 0
  mcPrevious = 1
 End Enum

Public Function GetQuestion(ByVal iQuestionType As mcGetQuestionTypes, Optional ByVal iMaxQuestions As Long = -1, Optional ByVal bResetList As Boolean = False) As Long
 Static sAlreadyAsked As String
 Static iMaxQ As Long
 Dim iQuestion As Long
 Dim Trash As String

 iQuestion = -1
 If bResetList Then sAlreadyAsked = vbNullString
 If iMaxQuestions <> -1 Then iMaxQ = iMaxQuestions
 If iMaxQ = -1 Then Exit Function
    
 Select Case iQuestionType
       
  Case mcNext
   If Len(sAlreadyAsked) Then
    Trash = Left$(sAlreadyAsked, Len(sAlreadyAsked) - 1)
    If InStr(Trash, ",") Then Trash = Mid$(Trash, InStrRev(Trash, ",") + 1)
     If CLng(Trash) <> iMaxQ Then iQuestion = CLng(Trash) + 1 ' Check for end of questions
    Else
     iQuestion = 1 ' Start at beginning
    End If
        
  Case mcPrevious
   If Len(sAlreadyAsked) Then
    Trash = Left$(sAlreadyAsked, Len(sAlreadyAsked) - 1)
    Trash = Mid$(Trash, InStrRev(Trash, ",") + 1)
    If CLng(Trash) > 0 Then iQuestion = CLng(Trash) - 1 ' At the beginning?
   End If
  End Select

  GetQuestion = iQuestion
  frmQuestions.lblQno.Caption = iQuestion & " out of " & iMaxQ
  sAlreadyAsked = sAlreadyAsked & iQuestion & ","
  
End Function
Public Sub ShowQuestion(ByVal iQuestion As Long)
 Dim oQuestion As mcQuestion

 Set oQuestion = oQuestions(iQuestion)
    
 With frmQuestions
  .chkAnswer(1).Visible = False: .chkAnswer(1).Value = vbUnchecked
  .chkAnswer(2).Visible = False: .chkAnswer(2).Value = vbUnchecked
  .chkAnswer(4).Visible = False: .chkAnswer(4).Value = vbUnchecked
  .chkAnswer(8).Visible = False: .chkAnswer(8).Value = vbUnchecked
  .chkAnswer(16).Visible = False: .chkAnswer(16).Value = vbUnchecked
  .optAnswer(1).Visible = False: .optAnswer(1).Value = False
  .optAnswer(2).Visible = False: .optAnswer(2).Value = False
  .optAnswer(4).Visible = False: .optAnswer(4).Value = False
  .optAnswer(8).Visible = False: .optAnswer(8).Value = False
  .optAnswer(16).Visible = False: .optAnswer(16).Value = False
        
  .lblQuestion.Caption = oQuestion.Question
  .lblQuestion.Tag = iQuestion
        
  If oQuestion.Multiple Then ' Show Option buttons
   If Len(oQuestion.Guess1) Then .chkAnswer(1).Visible = True: .chkAnswer(1).Caption = oQuestion.Guess1: .chkAnswer(1).Value = IIf(oQuestion.UserAnswer And 1, 1, 0)
   If Len(oQuestion.Guess2) Then .chkAnswer(2).Visible = True: .chkAnswer(2).Caption = oQuestion.Guess2: .chkAnswer(2).Value = IIf(oQuestion.UserAnswer And 2, 1, 0)
   If Len(oQuestion.Guess4) Then .chkAnswer(4).Visible = True: .chkAnswer(4).Caption = oQuestion.Guess4: .chkAnswer(4).Value = IIf(oQuestion.UserAnswer And 4, 1, 0)
   If Len(oQuestion.Guess8) Then .chkAnswer(8).Visible = True: .chkAnswer(8).Caption = oQuestion.Guess8: .chkAnswer(8).Value = IIf(oQuestion.UserAnswer And 8, 1, 0)
   If Len(oQuestion.Guess16) Then .chkAnswer(16).Visible = True: .chkAnswer(16).Caption = oQuestion.Guess16: .chkAnswer(16).Value = IIf(oQuestion.UserAnswer And 16, 1, 0)
  Else
   If Len(oQuestion.Guess1) Then .optAnswer(1).Visible = True: .optAnswer(1).Caption = oQuestion.Guess1
   If Len(oQuestion.Guess2) Then .optAnswer(2).Visible = True: .optAnswer(2).Caption = oQuestion.Guess2
   If Len(oQuestion.Guess4) Then .optAnswer(4).Visible = True: .optAnswer(4).Caption = oQuestion.Guess4
   If Len(oQuestion.Guess8) Then .optAnswer(8).Visible = True: .optAnswer(8).Caption = oQuestion.Guess8
   If Len(oQuestion.Guess16) Then .optAnswer(16).Visible = True: .optAnswer(16).Caption = oQuestion.Guess16
   If oQuestion.UserAnswer Then .optAnswer(oQuestion.UserAnswer).Value = 1
  End If
 End With
End Sub
