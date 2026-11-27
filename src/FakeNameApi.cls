VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeNameApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INameApi

' Test double: known names answer, everything else "no answer"; Offline
' simulates a stopped GCV application. Calls counts the requests.

Public Offline As Boolean
Public FailWithStatus As Long        ' e.g. 401 / 503: every request answers this status
Public Calls As Long

Private mAnswers As Object

Private Sub Class_Initialize()
    Set mAnswers = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Respond(ByVal arabicName As String, ByVal english As String)
    mAnswers(arabicName) = english
End Sub

Private Function INameApi_Transliterate(ByVal arabicName As String) As NameResult
    Dim result As NameResult
    Calls = Calls + 1
    Set result = New NameResult
    If Offline Then
        result.Outcome = noConnectionFailed
        result.ErrorText = "fake: offline"
    ElseIf FailWithStatus <> 0 Then
        result.HttpStatus = FailWithStatus
        Select Case FailWithStatus
            Case 401: result.Outcome = noUnauthorized
            Case 503: result.Outcome = noUnavailable
            Case Else: result.Outcome = noHttpError
        End Select
    ElseIf mAnswers.Exists(arabicName) Then
        result.Outcome = noOk
        result.HttpStatus = 200
        result.English = mAnswers(arabicName)
    Else
        result.Outcome = noNoAnswer
        result.HttpStatus = 200
    End If
    Set INameApi_Transliterate = result
End Function
