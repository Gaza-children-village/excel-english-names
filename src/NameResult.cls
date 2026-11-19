VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NameResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Answer of the name service for one Arabic name.

Public Enum NameOutcome
    noOk = 1
    noNoAnswer = 2            ' the service answered but gave no English name
    noHttpError = 3
    noConnectionFailed = 4
End Enum

Public Outcome As NameOutcome
Public English As String
Public HttpStatus As Long
Public ErrorText As String
