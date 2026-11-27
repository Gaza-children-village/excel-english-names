Attribute VB_Name = "EnglishNames"
Option Explicit

' ============================================================================
' GCV English Names add-in.
'
'   =TO_ENGLISH(A2)   English spelling (passport style) of the Arabic name in A2,
'                     from the name service of the GCV application.
'
' The result is a SUGGESTION to review. Ribbon "freeze reviewed names" turns
' the selected TO_ENGLISH formulas into fixed text, so a reviewed spelling
' never changes again and the file works without the add-in.
'
' Light on the PC: answers are kept for the Excel session (the same name is
' asked once); when the service cannot answer (not running, 401, 503 or
' another HTTP error), every TO_ENGLISH returns #N/A for 30 seconds without
' new requests. A worksheet function shows no messages: ribbon "try again"
' explains the last problem (asking /api/health), then asks again.
' Logs hold counts and failures, never names.
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Public Const DEFAULT_API_BASE As String = "http://127.0.0.1:5555/api/"
Private Const SETTINGS_APP As String = "GCV Tools"
Private Const SETTINGS_SECTION As String = "EnglishNames"
Private Const LOG_SOURCE As String = "english-names"
Private Const PAUSE_SECONDS As Long = 30
Private Const MB_ICONINFO As Long = &H40
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000

Private mApi As INameApi
Private mCache As Object               ' normalised Arabic name -> English
Private mPausedUntil As Date           ' no requests before this time (service down)
Private mRequests As Long
Private mTestMode As Boolean
Private mLastProblem As NameOutcome    ' 0 = none; kept for the next ribbon action
Private mLastStatus As Long

' ----------------------------------------------------------------------------
' The worksheet function
' ----------------------------------------------------------------------------

Public Function TO_ENGLISH(ByVal arabicName As Variant) As Variant
    Dim text As String
    Dim result As NameResult

    On Error GoTo Failed
    If IsObject(arabicName) Then
        If arabicName.Cells.Count <> 1 Then
            TO_ENGLISH = CVErr(xlErrValue)               ' one cell only
            Exit Function
        End If
        arabicName = arabicName.Value
    End If
    If IsError(arabicName) Then
        TO_ENGLISH = arabicName
        Exit Function
    End If
    text = NormaliseSpaces(CStr(arabicName))
    If Len(text) = 0 Then
        TO_ENGLISH = vbNullString
        Exit Function
    End If

    If Cache().Exists(text) Then
        TO_ENGLISH = Cache()(text)
        Exit Function
    End If
    If Now < mPausedUntil Then
        TO_ENGLISH = CVErr(xlErrNA)                      ' service down: no request
        Exit Function
    End If

    mRequests = mRequests + 1
    Set result = Api().Transliterate(text)
    Select Case result.Outcome
        Case noOk
            Cache()(text) = result.English
            mLastProblem = 0
            TO_ENGLISH = result.English
        Case noConnectionFailed, noUnauthorized, noUnavailable, noHttpError
            PauseAfter result
            TO_ENGLISH = CVErr(xlErrNA)
        Case Else
            TO_ENGLISH = CVErr(xlErrNA)                  ' no spelling for this name
    End Select
    Exit Function

Failed:
    LogError LOG_SOURCE, "TO_ENGLISH failed"
    TO_ENGLISH = CVErr(xlErrValue)
End Function

' Trims and collapses inner spaces (also non-breaking spaces).
Public Function NormaliseSpaces(ByVal text As String) As String
    text = Trim$(Replace(text, ChrW$(160), " "))
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    NormaliseSpaces = text
End Function

' UTF-8 percent-encoding for the query string.
Public Function UrlEncodeUtf8(ByVal text As String) As String
    Dim stream As Object
    Dim bytes() As Byte
    Dim i As Long
    Dim b As Byte
    Dim result As String

    If Len(text) = 0 Then Exit Function
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                            ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1                            ' binary
    stream.Position = 3                        ' skip the UTF-8 byte order mark
    bytes = stream.Read
    stream.Close
    For i = LBound(bytes) To UBound(bytes)
        b = bytes(i)
        If (b >= 48 And b <= 57) Or (b >= 65 And b <= 90) Or (b >= 97 And b <= 122) _
           Or b = 45 Or b = 46 Or b = 95 Or b = 126 Then
            result = result & Chr$(b)
        Else
            result = result & "%" & Right$("0" & Hex$(b), 2)
        End If
    Next i
    UrlEncodeUtf8 = result
End Function

' ----------------------------------------------------------------------------
' Add-in life cycle and ribbon
' ----------------------------------------------------------------------------

' The service cannot answer now: no requests for PAUSE_SECONDS, one log line,
' and the reason kept for the next ribbon action.
Private Sub PauseAfter(ByVal result As NameResult)
    mPausedUntil = Now + TimeSerial(0, 0, PAUSE_SECONDS)
    mLastProblem = result.Outcome
    mLastStatus = result.HttpStatus
    Select Case result.Outcome
        Case noConnectionFailed
            LogWarn LOG_SOURCE, "Name service unreachable (" & result.ErrorText & "); paused for " & _
                                PAUSE_SECONDS & " s after " & mRequests & " request(s)."
        Case noUnauthorized
            LogWarn LOG_SOURCE, "Name service refused the request (HTTP 401): the GCV application must be " & _
                                "opened or updated; paused for " & PAUSE_SECONDS & " s."
        Case noUnavailable
            LogWarn LOG_SOURCE, "Name service not ready (HTTP 503); paused for " & PAUSE_SECONDS & " s."
        Case Else
            LogWarn LOG_SOURCE, "Name service returned HTTP " & result.HttpStatus & "; paused for " & _
                                PAUSE_SECONDS & " s."
    End Select
End Sub

Public Function LastProblem() As NameOutcome
    LastProblem = mLastProblem
End Function

' Message for a problem; healthJson = answer of /api/health ("" = no answer).
Public Function ProblemMessage(ByVal problem As NameOutcome, ByVal httpStatus As Long, _
                               ByVal healthJson As String) As String
    Dim namesState As String
    Dim appVersion As String

    ReadHealthParts healthJson, namesState, appVersion
    Select Case problem
        Case noUnauthorized: ProblemMessage = T_Unauthorized()
        Case noUnavailable: ProblemMessage = T_NotReady()
        Case noConnectionFailed
            If namesState = "unavailable" Then ProblemMessage = T_NotReady() Else ProblemMessage = T_Offline()
        Case noHttpError: ProblemMessage = T_HttpError(httpStatus)
        Case Else: Exit Function
    End Select
    If Len(appVersion) > 0 Then ProblemMessage = ProblemMessage & vbCrLf & T_AppVersion(appVersion)
End Function

Private Sub ReadHealthParts(ByVal healthJson As String, ByRef namesState As String, ByRef appVersion As String)
    Dim answer As Object
    If Len(healthJson) = 0 Then Exit Sub
    On Error Resume Next                         ' a malformed answer only loses the detail
    Set answer = JsonParse(healthJson)
    namesState = LCase$(Trim$(JsonStr(answer, "names")))
    appVersion = JsonStr(answer, "version")
    On Error GoTo 0
End Sub

Public Function T_Unauthorized() As String
    T_Unauthorized = FromCodePoints("0631 0641 0636 0020 062A 0637 0628 064A 0642 0020 0047 0043 0056 0020 0637 0644 0628 0020 0627 0644 0623 0633 0645 0627 0621 0020 0028 0645 0641 062A 0627 062D 0020 0627 0644 0627 062A 0635 0627 0644 0020 063A 064A 0631 0020 0645 0642 0628 0648 0644 0029 002E") & vbCrLf & _
                     FromCodePoints("0627 0641 062A 062D 0020 062A 0637 0628 064A 0642 0020 0047 0043 0056 0020 0623 0648 0020 062D 062F 0651 062B 0647 060C 0020 062B 0645 0020 0627 0636 063A 0637 0020 00AB 0625 0639 0627 062F 0629 0020 0627 0644 0645 062D 0627 0648 0644 0629 00BB 002E")
End Function

Public Function T_NotReady() As String
    T_NotReady = FromCodePoints("062E 062F 0645 0629 0020 0627 0644 0623 0633 0645 0627 0621 0020 063A 064A 0631 0020 062C 0627 0647 0632 0629 0020 0641 064A 0020 062A 0637 0628 064A 0642 0020 0047 0043 0056 0020 0628 0639 062F 002E 0020 062D 0627 0648 0644 0020 0644 0627 062D 0642 0627 064B 0020 0628 0632 0631 0020 00AB 0625 0639 0627 062F 0629 0020 0627 0644 0645 062D 0627 0648 0644 0629 00BB 002E")
End Function

Public Function T_Offline() As String
    T_Offline = FromCodePoints("062A 0639 0630 0631 0020 0627 0644 0627 062A 0635 0627 0644 0020 0628 062A 0637 0628 064A 0642 0020 0047 0043 0056 002E 0020 062A 0623 0643 062F 0020 0645 0646 0020 062A 0634 063A 064A 0644 0647 060C 0020 062B 0645 0020 0627 0636 063A 0637 0020 00AB 0625 0639 0627 062F 0629 0020 0627 0644 0645 062D 0627 0648 0644 0629 00BB 002E")
End Function

Public Function T_HttpError(ByVal httpStatus As Long) As String
    T_HttpError = FromCodePoints("0623 0639 0627 062F 0020 062A 0637 0628 064A 0642 0020 0047 0043 0056 0020 062E 0637 0623 0020 0028 0048 0054 0054 0050 0020") & httpStatus & FromCodePoints("0029 002E 0020 062D 0627 0648 0644 0020 0644 0627 062D 0642 0627 064B 0020 0623 0648 0020 0627 0633 062A 062E 062F 0645 0020 00AB 0627 0644 0625 0628 0644 0627 063A 0020 0639 0646 0020 0645 0634 0643 0644 0629 00BB 002E")
End Function

Public Function T_AppVersion(ByVal appVersion As String) As String
    T_AppVersion = FromCodePoints("0625 0635 062F 0627 0631 0020 062A 0637 0628 064A 0642 0020 0047 0043 0056 003A 0020") & appVersion
End Function

Public Sub AddinOpened()
    On Error Resume Next                       ' start-up must never show a VBA error
    LogInfo LOG_SOURCE, GcvAddin.ADDIN_NAME & " " & GcvAddin.ADDIN_VERSION & " opened."
    ' Description in the Insert Function dialog (may be refused for a read-only add-in).
    Application.MacroOptions Macro:="TO_ENGLISH", _
        Description:="English spelling (passport style) of an Arabic name - a suggestion to review.", _
        Category:="GCV Tools", ArgumentDescriptions:=Array("Cell with the Arabic name")
End Sub

Public Sub AddinClosing()
End Sub

Public Sub AddinSelfCheck()
    NormaliseSpaces vbNullString
    UrlEncodeUtf8 "a"
    GcvLocalApi.HealthUrl vbNullString
End Sub

' Turns the selected TO_ENGLISH formulas with a result into fixed text.
Public Sub Ribbon_FreezeNames(control As IRibbonControl)
    Dim count As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    count = FreezeNames(Selection)
    LogInfo LOG_SOURCE, "Frozen: " & count & " name(s)."
    Notify FromCodePoints("062A 0645 0020 062A 062B 0628 064A 062A 0020") & count & FromCodePoints("0020 0627 0633 0645 0627 064B 0020 0643 0646 0635 002E 0020 0627 0644 062E 0644 0627 064A 0627 0020 0627 0644 062A 064A 0020 0641 064A 0647 0627 0020 0023 004E 002F 0041 0020 0644 0645 0020 062A 062A 063A 064A 0631 002E")
End Sub

' Explains the last problem (if any), forgets it and asks the service again
' for the selected cells.
Public Sub Ribbon_TryAgain(control As IRibbonControl)
    Dim problem As NameOutcome
    problem = mLastProblem
    mPausedUntil = 0
    mLastProblem = 0
    LogInfo LOG_SOURCE, "Try again requested."
    If problem <> 0 And Not mTestMode Then
        Notify ProblemMessage(problem, mLastStatus, GcvLocalApi.ReadHealth(GcvLocalApi.HealthUrl(ApiBaseUrl())))
    End If
    If TypeName(Selection) = "Range" Then
        Selection.Dirty
        Selection.Calculate
    End If
End Sub

Public Function FreezeNames(ByVal target As Range) As Long
    Dim cell As Range
    Dim area As Range

    Set area = Intersect(target, target.Worksheet.UsedRange)
    If area Is Nothing Then Exit Function
    For Each cell In area.Cells
        If cell.HasFormula Then
            If InStr(1, cell.Formula, "TO_ENGLISH(", vbTextCompare) > 0 And Not IsError(cell.Value) Then
                cell.Value = cell.Value
                FreezeNames = FreezeNames + 1
            End If
        End If
    Next cell
End Function

' ----------------------------------------------------------------------------
' Service, cache, tests
' ----------------------------------------------------------------------------

' Per PC (registry): HKCU\Software\VB and VBA Program Settings\GCV Tools\EnglishNames\ApiBaseUrl
Public Function ApiBaseUrl() As String
    ApiBaseUrl = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "ApiBaseUrl", DEFAULT_API_BASE)
End Function

Private Function Api() As INameApi
    Dim httpApi As HttpNameApi
    If mApi Is Nothing Then
        Set httpApi = New HttpNameApi
        httpApi.Init ApiBaseUrl()
        Set mApi = httpApi
    End If
    Set Api = mApi
End Function

Private Function Cache() As Object
    If mCache Is Nothing Then Set mCache = CreateObject("Scripting.Dictionary")
    Set Cache = mCache
End Function

Public Sub UseTestApi(ByVal fake As INameApi)
    Set mApi = fake
    Set mCache = Nothing
    mPausedUntil = 0
    mRequests = 0
    mLastProblem = 0
    mLastStatus = 0
    mTestMode = True
End Sub

Public Sub ResetForTests()
    Set mApi = Nothing
    Set mCache = Nothing
    mPausedUntil = 0
    mRequests = 0
    mLastProblem = 0
    mLastStatus = 0
    mTestMode = False
End Sub

Private Sub Notify(ByVal text As String)
    If mTestMode Then Exit Sub
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(FromCodePoints("0627 0644 0623 0633 0645 0627 0621 0020 0627 0644 0625 0646 062C 0644 064A 0632 064A 0629")), MB_ICONINFO Or MB_RIGHT Or MB_RTLREADING
End Sub

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function
