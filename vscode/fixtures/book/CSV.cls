VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Read one column.
' @param index Column number.
' @returns The cell text.
Public Function GetColumn(index As Long) As String
    GetColumn = CStr(index)
End Function
