Attribute VB_Name = "txt"
Option Explicit

Function txt_write(ByVal strText As String, Optional ByVal strOutputFileName As String, Optional ByVal appendToOutput As Boolean = True) As String
  txt_write = "!!!Error"
On Error GoTo lbl_exit
  If strOutputFileName = "" Then
    strOutputFileName = path_dropbox() & txt_yyyymmddhhmmss() & ".txt"
  End If
  Shell ("touch '" & strOutputFileName & "'")
  Open strOutputFileName For Append As #1
    Print #1, strText
  Close #1
  txt_write = strOutputFileName
lbl_exit:
End Function

Function txt_yyyymmddhhmmss() As String
  txt_yyyymmddhhmmss = Format(Now(), "yyyymmddhhmmss")
End Function
