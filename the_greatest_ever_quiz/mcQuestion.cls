VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mcQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Question"
'The original author of this code is Dave Applegate
Option Explicit
Private pQuestionID As Long
Private pQuestion As String
Private pGuess1 As String
Private pGuess2 As String
Private pGuess4 As String
Private pGuess8 As String
Private pGuess16 As String
Private pMultiple As Boolean
Private pAnswer As Long
Private pUserAnswer As Long
Private pIndex As Long

Public Property Let Answer(ByVal vData As Long)
 pAnswer = vData
End Property


Public Property Get Answer() As Long
 Answer = pAnswer
End Property
Public Property Let Index(ByVal vData As Long)
 pIndex = vData
End Property
Public Property Get Index() As Long
 Index = pIndex
End Property
Public Property Let Multiple(ByVal vData As Boolean)
 pMultiple = vData
End Property
Public Property Get Multiple() As Boolean
 Multiple = pMultiple
End Property
Public Property Let Guess16(ByVal vData As String)
 pGuess16 = vData
End Property
Public Property Get Guess16() As String
 Guess16 = pGuess16
End Property
Public Property Let Guess8(ByVal vData As String)
 pGuess8 = vData
End Property
Public Property Get Guess8() As String
 Guess8 = pGuess8
End Property
Public Property Let Guess4(ByVal vData As String)
 pGuess4 = vData
End Property
Public Property Get Guess4() As String
 Guess4 = pGuess4
End Property
Public Property Let Guess2(ByVal vData As String)
 pGuess2 = vData
End Property
Public Property Get Guess2() As String
 Guess2 = pGuess2
End Property
Public Property Let Guess1(ByVal vData As String)
 pGuess1 = vData
End Property
Public Property Get Guess1() As String
 Guess1 = pGuess1
End Property
Public Property Let Question(ByVal vData As String)
 pQuestion = vData
End Property
Public Property Get Question() As String
Attribute Question.VB_UserMemId = 0
 Question = pQuestion
End Property
Public Property Get QuestionID() As Long
 QuestionID = pQuestionID
End Property
Public Property Let QuestionID(ByVal vData As Long)
 pQuestionID = vData
End Property
Public Property Let UserAnswer(ByVal vData As Long)
 pUserAnswer = vData
End Property
Public Property Get UserAnswer() As Long
 UserAnswer = pUserAnswer
End Property
