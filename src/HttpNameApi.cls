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

' Name service of the GCV application (PROPOSED contract - the endpoint does
' not exist yet):
'   GET <base>transliterate?name=<UTF-8, URL-encoded Arabic name>
'   200 {"english": "Mohammed Ahmed Khaled Abu Shamala", "needsReview": true}
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
    http.SetRequestHeader "X-GCV-Excel-Addin", GcvAddin.ADDIN_ID & "/" & GcvAddin.ADDIN_VERSION
    http.Send
    On Error GoTo 0

    result.HttpStatus = http.Status
    If http.Status = 200 Then
        Set answer = JsonParse(http.ResponseText)
        result.English = Trim$(JsonStr(answer, "english"))
        result.Outcome = IIf(Len(result.English) > 0, noOk, noNoAnswer)
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
