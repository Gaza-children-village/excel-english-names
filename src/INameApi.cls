VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INameApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Boundary to the name service. HttpNameApi is the real one; FakeNameApi is
' used by the tests.

Public Function Transliterate(ByVal arabicName As String) As NameResult
End Function
