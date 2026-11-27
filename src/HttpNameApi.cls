VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpNameApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INameApi

' Name service of the GCV application (desktop app v1.10.0 and later,
' contract in docs/local-api.md of the app):
'   GET <base>transliterate?name=<UTF-8, URL-encoded Arabic name>
'   200 {"english": "Mohammed Ahmed Khaled Abu Shamala", "needsReview": true}
'       ("english" empty: no spelling possible)
'   401 key refused (open or update the app), 503 name service not ready
' Short time-outs: a formula must not freeze Excel when the app is not running.

Private Const CONNECT_TIMEOUT_MS As Long = 1500
Private Const ANSWER_TIMEOUT_MS As Long = 4000

Private mBaseUrl As String

Public Sub Init(ByVal baseUrl As String)
    mBaseUrl = baseUrl
    If Right$(mBaseUrl, 1) <> "/" Then mBaseUrl = mBaseUrl & "/"
End Sub

Private Function INameApi_Transliterate(ByVal arabicName As String) As NameResult
    Dim result As NameResult
    Dim http As Object
    Dim answer As Object

    Set result = New NameResult
    On Error GoTo Offline
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts CONNECT_TIMEOUT_MS, CONNECT_TIMEOUT_MS, ANSWER_TIMEOUT_MS, ANSWER_TIMEOUT_MS
    http.Open "GET", mBaseUrl & "transliterate?name=" & UrlEncodeUtf8(arabicName), False
    http.SetRequestHeader "Accept", "application/json"
    GcvLocalApi.AddHeaders http                     ' add-in version and the app's key
    http.Send
    On Error GoTo 0

    result.HttpStatus = http.Status
    If http.Status = 200 Then
        Set answer = JsonParse(http.ResponseText)
        result.English = Trim$(JsonStr(answer, "english"))
        result.Outcome = IIf(Len(result.English) > 0, noOk, noNoAnswer)
    ElseIf http.Status = GcvLocalApi.HTTP_UNAUTHORIZED Then
        result.Outcome = noUnauthorized
    ElseIf http.Status = GcvLocalApi.HTTP_UNAVAILABLE Then
        result.Outcome = noUnavailable
    Else
        result.Outcome = noHttpError
    End If
    Set INameApi_Transliterate = result
    Exit Function

Offline:
    result.Outcome = noConnectionFailed
    result.ErrorText = Err.Description
    Set INameApi_Transliterate = result
End Function
