Attribute VB_Name = "TestEnglishNames"
Option Explicit

' Tests for GCV English Names. Run RunEnglishNamesTests from the VBA editor
' (F5) in the DEV copy. Uses FakeNameApi (no network) and a temporary
' workbook closed without saving. Do NOT ship this module or FakeNameApi.

Private passedCount As Long
Private failedCount As Long
Private currentTest As String

Public Sub RunEnglishNamesTests()
    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "English names tests"
    On Error GoTo Crash

    Test_NormaliseSpaces
    Test_UrlEncode_SendsArabicAsUtf8
    Test_ToEnglish_ReturnsTheSpellingAndAsksOnce
    Test_ToEnglish_EmptyAndErrorInputs
    Test_ToEnglish_UnknownNameIsNA
    Test_ToEnglish_PausesRequestsWhenServiceIsDown
    Test_ToEnglish_RefusesSeveralCells
    Test_FreezeNames_TurnsResultsIntoText

Finish:
    EnglishNames.ResetForTests
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "English names tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

Private Sub Test_NormaliseSpaces()
    StartTest "Spaces are trimmed and collapsed"
    AssertEqual "a b c", NormaliseSpaces("  a   b" & ChrW$(160) & "c "), "collapsed"
    AssertEqual "", NormaliseSpaces("   "), "empty"
End Sub

Private Sub Test_UrlEncode_SendsArabicAsUtf8()
    StartTest "Arabic names are sent as UTF-8"
    AssertEqual "%D9%85%D8%AD%D9%85%D8%AF%20%D8%B9%D9%84%D9%8A", UrlEncodeUtf8(FromCodePoints("0645 062D 0645 062F 0020 0639 0644 064A")), "Arabic"
    AssertEqual "Abu-Shamala_1.x~", UrlEncodeUtf8("Abu-Shamala_1.x~"), "unreserved kept"
End Sub

Private Sub Test_ToEnglish_ReturnsTheSpellingAndAsksOnce()
    Dim fake As FakeNameApi
    StartTest "Known name: spelling returned, asked once per session"
    Set fake = Prepared()
    AssertEqual "Mohammed Ali", TO_ENGLISH(FromCodePoints("0645 062D 0645 062F 0020 0639 0644 064A")), "spelling"
    AssertEqual "Mohammed Ali", TO_ENGLISH("  " & FromCodePoints("0645 062D 0645 062F 0020 0020 0639 0644 064A") & " "), "same name, other spaces"
    AssertEqual 1, fake.Calls, "one request"
End Sub

Private Sub Test_ToEnglish_EmptyAndErrorInputs()
    Dim fake As FakeNameApi
    StartTest "Empty input gives empty text; an error stays an error"
    Set fake = Prepared()
    AssertEqual "", TO_ENGLISH(""), "empty"
    AssertEqual True, IsError(TO_ENGLISH(CVErr(xlErrDiv0))), "error passed on"
    AssertEqual 0, fake.Calls, "no request"
End Sub

Private Sub Test_ToEnglish_UnknownNameIsNA()
    StartTest "A name without spelling gives #N/A"
    Prepared
    AssertEqual CStr(CVErr(xlErrNA)), CStr(TO_ENGLISH(FromCodePoints("0627 0633 0645 0020 063A 064A 0631 0020 0645 0639 0631 0648 0641"))), "#N/A"
End Sub

Private Sub Test_ToEnglish_PausesRequestsWhenServiceIsDown()
    Dim fake As FakeNameApi
    StartTest "Service down: #N/A and no more requests for a while"
    Set fake = Prepared()
    fake.Offline = True
    AssertEqual CStr(CVErr(xlErrNA)), CStr(TO_ENGLISH(FromCodePoints("0645 062D 0645 062F 0020 0639 0644 064A"))), "first"
    AssertEqual CStr(CVErr(xlErrNA)), CStr(TO_ENGLISH(FromCodePoints("0623 062D 0645 062F"))), "second"
    AssertEqual 1, fake.Calls, "paused after one failure"
End Sub

Private Sub Test_ToEnglish_RefusesSeveralCells()
    Dim wb As Workbook
    StartTest "A range of several cells gives #VALUE!"
    Prepared
    Set wb = Workbooks.Add(xlWBATWorksheet)
    AssertEqual CStr(CVErr(xlErrValue)), CStr(TO_ENGLISH(wb.Worksheets(1).Range("A1:A2"))), "#VALUE!"
    wb.Close SaveChanges:=False
End Sub

Private Sub Test_FreezeNames_TurnsResultsIntoText()
    Dim wb As Workbook
    Dim ws As Worksheet
    StartTest "Freeze: results become text, #N/A cells keep their formula"
    Prepared
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = FromCodePoints("0645 062D 0645 062F 0020 0639 0644 064A")
    ws.Range("A2").Value = FromCodePoints("0627 0633 0645 0020 063A 064A 0631 0020 0645 0639 0631 0648 0641")
    ws.Range("B1").Formula = "=TO_ENGLISH(A1)"
    ws.Range("B2").Formula = "=TO_ENGLISH(A2)"
    ws.Range("B1:B2").Calculate
    AssertEqual 1, FreezeNames(ws.Range("B1:B2")), "one frozen"
    AssertEqual False, ws.Range("B1").HasFormula, "B1 is text"
    AssertEqual "Mohammed Ali", ws.Range("B1").Value, "B1 value"
    AssertEqual True, ws.Range("B2").HasFormula, "B2 keeps its formula"
    wb.Close SaveChanges:=False
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function Prepared() As FakeNameApi
    Dim fake As FakeNameApi
    Set fake = New FakeNameApi
    fake.Respond FromCodePoints("0645 062D 0645 062F 0020 0639 0644 064A"), "Mohammed Ali"
    EnglishNames.UseTestApi fake
    Set Prepared = fake
End Function

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    Dim same As Boolean
    If IsError(expected) Or IsError(actual) Then
        same = (CStr(expected) = CStr(actual))
    ElseIf VarType(expected) = vbBoolean Or VarType(actual) = vbBoolean Then
        same = (CStr(expected) = CStr(actual))
    ElseIf IsNumeric(expected) And IsNumeric(actual) And VarType(expected) <> vbString And VarType(actual) <> vbString Then
        same = (CDbl(expected) = CDbl(actual))
    Else
        same = (CStr(expected) = CStr(actual))
    End If
    If same Then
        passedCount = passedCount + 1
    Else
        failedCount = failedCount + 1
        Debug.Print "FAIL: " & currentTest & " / " & label & " - expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
    End If
End Sub

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function
