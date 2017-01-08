Attribute VB_Name = "path"
Option Explicit

Function path_dropbox() As String
  path_dropbox = Environ("DROPBOX") & "\"
End Function


