Attribute VB_Name = "A"
Option Explicit
Option Base 1

Function a_ndim(x) As Byte '0-255
'returns the dimension count of array x
    'non-array -> all deemed as scalar
    If Not IsArray(x) Then Exit Function
    
    'array
    Dim res As Byte: res = 0
    Dim isXlRange As Boolean: isXlRange = (TypeName(x) = "Range")
On Error GoTo lbl_exit
    For res = 0 To 255
        If isXlRange Then
            If UBound(x.Value2, res + 2) - LBound(x.Value2, res + 2) >= 0 Then
            End If
        Else
            If UBound(x, res + 2) - LBound(x, res + 2) >= 0 Then
            End If
        End If
    Next res
lbl_exit:
    a_ndim = res + 1  'makesure it errors out when >255
End Function

Function a_dims(x)
    Dim ndim As Byte: ndim = a_ndim(x)
    If 0 = ndim Then a_dims = Empty: Exit Function
    Dim isXlRange As Boolean: isXlRange = (TypeName(x) = "Range")
    Dim res: ReDim res(1 To ndim)
    Dim i As Byte
    For i = 1 To ndim
      If isXlRange Then
        res(i) = UBound(x.Value2, i) - LBound(x.Value2, i) + 1
      Else
        res(i) = UBound(x, i) - LBound(x, i) + 1
      End If
    Next i
    a_dims = res
End Function

Function a_dims2(x)
    Dim ndim As Byte: ndim = a_ndim(x)
    If 0 = ndim Then a_dims2 = Empty: Exit Function
    Dim isXlRange As Boolean: isXlRange = (TypeName(x) = "Range")
    Dim res: ReDim res(1 To ndim, 1 To 2)
    Dim i As Byte
    For i = 1 To ndim
      If isXlRange Then
        res(i, 1) = LBound(x.Value2, i): res(i, 2) = UBound(x.Value2, i)
      Else
        res(i, 1) = LBound(x, i): res(i, 2) = UBound(x, i)
      End If
    Next i
    a_dims2 = res
End Function
