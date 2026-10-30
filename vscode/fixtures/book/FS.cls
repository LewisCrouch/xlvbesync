VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Join path fragments.
' @param parts Path pieces.
' @returns The combined path.
Public Function JoinPaths(parts As Variant) As String
    JoinPaths = ""
End Function

Private Const cMaxFilePathLength As Long = 218

' @param filePath Path to test.
Public Function IsWithinFilePathLimit(filePath As String) As Boolean
    IsWithinFilePathLimit = Len(filePath) <= cMaxFilePathLength
    MaxFilePathLength = cMaxFilePathLength
End Function
