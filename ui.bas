Attribute VB_Name = "UI"
Option Explicit

Function yy(ByVal strYcolName As String, ydata)
  yy = "y:" & strYcolName
End Function

Function ee(ByVal strErrFunc As String, ypredicted, yobserved) As String
lbl_validate_same_dims:
  Dim ndim As Byte: ndim = a_ndim(ypredicted)
  If ndim <> a_ndim(yobserved) Then GoTo lbl_error
  Dim i As Byte
  Select Case ndim
    Case 0: 'pass
    Case 1: If UBound(ypredicted) - LBound(ypredicted) <> UBound(yobserved) - LBound(yobserved) Then GoTo lbl_error
    Case 2:
      Dim dims1: dims1 = a_dims(ypredicted): If Application.Product(dims1) <> Application.Max(dims1) Then GoTo lbl_error
      Dim dims2: dims2 = a_dims(yobserved)
      For i = LBound(dims1) To UBound(dims1)
        If dims1(i) <> dims2(i) Then GoTo lbl_error
      Next i
    Case Else: GoTo lbl_error
  End Select: GoTo lbl_return
  
lbl_error: strErrFunc = CVErr(xlErrValue)

lbl_return:
  Select Case LCase(strErrFunc)
    Case "sse": ee = "L2"
    Case "xen": ee = "xentropy"
    Case Else: ee = strErrFunc
  End Select
  ee = "e:" & ee
End Function

Function oo(ByVal strActFunc As String, ParamArray inputNeurons() As Variant)
  Select Case LCase(strActFunc)
    Case "l", "logit", "logistic": oo = "logit"
    Case "mlogit": oo = "mlogit"
    Case "id": oo = "id"
    Case "lin", "linear": oo = "linear"
    Case Else: oo = CVErr(xlErrNA): Exit Function
  End Select
  oo = "o:" & oo
End Function

Function nn(ByVal strActFunc As String, ParamArray inputNeurons() As Variant)
  Select Case LCase(strActFunc)
    Case "1": nn = "1"
    Case "l", "logit", "logistic": nn = "logit"
    Case Else: nn = CVErr(xlErrNA)
  End Select
End Function

Function ii(ByVal strInputColName As String, indata) As String
  ii = "i:" & strInputColName
End Function

Sub nn_select(Optional ByVal wsName As String = "")
  'shtPrep.Range(shtPrep.Shapes(Application.Caller).TextFrame2.TextRange.Characters.Text).Select
  Dim ws As Worksheet
  If wsName = "" Then Set ws = shtPrep Else Set ws = ActiveSheet
  ws.Range(Replace(Application.Caller, "Oval_", "")).Select
End Sub

Sub createNeuron(ByVal x As Double, ByVal y As Double, ByVal displayText As String, Optional ByVal nameText As String, Optional ws As Worksheet)
    Application.EnableEvents = False
    If ws Is Nothing Then Set ws = shtPrep
    Dim sh As Shape: Set sh = ws.Shapes.AddShape(msoShapeOval, x, y, 30, 30)
    If nameText = "" Then nameText = displayText
    sh.Name = "Oval_" & nameText
    With sh.TextFrame2
      .MarginLeft = 2.5
      .MarginRight = 2.5
      .MarginTop = 3
      .MarginBottom = 3
      .WordWrap = msoFalse
      .VerticalAnchor = msoAnchorMiddle
      .HorizontalAnchor = msoAnchorCenter
      .TextRange.Characters.Text = displayText
    End With
    sh.OnAction = "'nn_select """ & ws.Name & """'"
End Sub

Sub clearShapes(ws As Worksheet)
  Application.EnableEvents = False
  Dim x
  For Each x In ws.Shapes
    If Not x.Name Like "cbtn*" Then x.Delete
  Next x
  Application.EnableEvents = True
End Sub

Sub addConnector(ByVal strStartCell As String, ByVal strEndCell As String, Optional ws As Worksheet, Optional ByVal strName As String)
  If ws Is Nothing Then Set ws = shtPrep
  Debug.Print strStartCell, strEndCell
  Dim sh As Shape
  Set sh = ws.Shapes.addConnector(msoConnectorStraight, 0, 0, 0, 0)
  If strName <> "" Then sh.Name = "Conn_" & Replace(strName, "$", "")
  sh.Line.EndArrowheadStyle = msoArrowheadTriangle
  sh.ConnectorFormat.BeginConnect ws.Shapes("Oval_" & Replace(strStartCell, "$", "")), 7
  sh.ConnectorFormat.EndConnect ws.Shapes("Oval_" & Replace(strEndCell, "$", "")), 3
  'sh.RerouteConnections
  With sh.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 192, 0)
    .Transparency = 0
  End With
End Sub

Sub emitConnector(ByVal neuron As Range, Optional ws As Worksheet)
  Dim r As Range
  On Error GoTo e
  For Each r In neuron.DirectDependents.Cells
    Call addConnector(neuron.AddressLocal, r.AddressLocal, ws)
  Next r
e:
End Sub

Function logit(ByVal x)
    'logit = Application.Tanh(x / 2) / 2 + 0.5
    logit = 1 / (1 + Exp(-x))
End Function

Function pre(neuron As Range) As Range
  Dim i&, j&, k&, M(1 To 3) As Long, s As String: s = neuron.FormulaLocal
  i = InStr(s, "("): j = InStr(s, ")")
  If Not (i > 0 And j > i) Then Exit Function
  For k = i + 1 To j - 1
    Select Case Mid(s, k, 1)
      Case "(": M(1) = M(1) + 1
      Case ")": M(1) = M(1) - 1
      Case "[": M(2) = M(2) + 1
      Case "]": M(2) = M(2) - 1
      Case "{": M(3) = M(3) + 1
      Case "}": M(3) = M(3) - 1
      Case ",":
        If M(1) = 0 And M(2) = 0 And M(3) = 0 Then
          'not inside any braces=> next arg begins
          Set pre = neuron.Worksheet.Range(Trim(Mid(s, k + 1, j - 1 - k)))
          Exit Function
        End If
    End Select
  Next k
End Function

Function dep(neuron As Range) As Range
  Set dep = neuron.DirectDependents
End Function


Sub colorRangeBorder(W As Range, ByVal indexThemeColor As Integer)
    W.Borders(xlDiagonalDown).LineStyle = xlNone
    W.Borders(xlDiagonalUp).LineStyle = xlNone
    With W.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = indexThemeColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With W.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = indexThemeColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With W.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = indexThemeColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With W.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = indexThemeColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    W.Borders(xlInsideVertical).LineStyle = xlNone
    W.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub underscoreFormulaCells(W As Range)
      W.FormatConditions.Add Type:=xlExpression, Formula1:="=ISFORMULA(" & Replace(W.Cells(1, 1).AddressLocal, "$", "") & ")"
    W.FormatConditions(W.FormatConditions.Count).SetFirstPriority
    With W.FormatConditions(1).Font
        .Underline = xlUnderlineStyleSingle
        .TintAndShade = 0
    End With
    W.FormatConditions(1).StopIfTrue = False
End Sub



Sub imputeNoise(ws As Worksheet)
On Error GoTo 0
  Dim n As Byte, r As Range
  n = getNumLayers(ws)
  While n >= 1
    For Each r In ws.Range("W_" & n).Cells
      If Trim(r.Value) <> "" And Not r.hasFormula And IsNumeric(r) Then
        r.Value = r.Value + Application.Max(Abs(r.Value), 0.0001) * rnorm1() / 20
      End If
    Next r
    n = n - 1
  Wend
  ws.Calculate
End Sub

Sub trimWeiMats(ws As Worksheet)
On Error GoTo 0
  Dim n As Byte, r As Range, W As Double, loss As Double, tol As Double
  tol = CDbl(InputBox("Enter maximum change in totloss to trim:", "Specify Trim Tolerance", 0.0005))
  n = getNumLayers(ws)
  While n >= 1
    For Each r In ws.Range("W_" & n).Cells
      If Trim(r.Formula) <> "" And IsNumeric(r.Formula) Then
        W = r.Value
        loss = ws.Range("totloss").Value
        r.Value = 0: ws.Calculate
        If Abs(loss - ws.Range("totloss").Value) < tol Then
          r.FormulaLocal = "=0"
        Else
          r.Value = W
        End If
      End If
    Next r
    n = n - 1
  Wend
  ws.Calculate
End Sub

Sub shrinkWeiMats(ws As Worksheet)
On Error GoTo 0
  Dim n As Byte, r As Range, W As Double, loss As Double, scalar As Double
  scalar = CDbl(InputBox("Enter shrinkage scalar:", "Rescale weights", 0.5))
  n = getNumLayers(ws)
  While n >= 1
    For Each r In ws.Range("W_" & n).Cells
      If Trim(r.Formula) <> "" And IsNumeric(r.Formula) Then
        r.Value = scalar * r.Value
      End If
    Next r
    n = n - 1
  Wend
  ws.Calculate
End Sub

Function getNumLayers(ws As Worksheet) As Byte
  Dim n As Byte
  On Error GoTo lbl_1
  While True
    If Len(ws.Range("W_" & (n + 1)).AddressLocal) > 0 Then
    End If
    n = n + 1
  Wend
lbl_1:
  getNumLayers = n
End Function

Function hasFormula(r As Range) As Boolean
  hasFormula = r.hasFormula
End Function


Sub plotNetOnInstanceSheet(ws As Worksheet)
  Call clearShapes(ws)
  Dim i&, j&, k&, M&, n&, nlayers&: nlayers = getNumLayers(ws)
  
  Dim r As Range, c As Range, W As Range
  For M = 1 To nlayers
    k = 1
    For Each r In ws.Range("N_" & M - 1).Cells
      createNeuron M * 120, k * 45, r.Value, Replace(r.Address, "$", ""), ws
      k = k + 1
    Next r
    If M > 1 Then
      Set W = ws.Range("W_" & M - 1)
      For i = 1 To W.Rows.Count
        For j = 1 To W.Columns.Count
          If W.Cells(i, j).Formula <> "=0" And CStr(ws.Range("N_" & M - 1).Cells(j, 1).Value) <> "1" Then
            addConnector ws.Range("N_" & M - 2).Cells(i, 1).AddressLocal, ws.Range("N_" & M - 1).Cells(j, 1).AddressLocal, ws, W.Cells(i, j).AddressLocal
          End If
        Next j
      Next i
    End If
  Next M
lbl_yhat:
  k = 1
  For Each r In ws.Range("yhat").Columns(1).Offset(0, -1).Cells
    createNeuron (nlayers + 1) * 120, k * 45, r.Value, Replace(r.Address, "$", ""), ws
    k = k + 1
  Next r
  Set W = ws.Range("W_" & nlayers)
  Dim yhatNamesRange As Range: Set yhatNamesRange = ws.Range("yhat").Columns(1).Offset(0, -1)
  For i = 1 To W.Rows.Count
    For j = 1 To W.Columns.Count
      If W.Cells(i, j).Formula <> "=0" And CStr(yhatNamesRange.Cells(j, 1).Value) <> "1" Then
        addConnector ws.Range("N_" & nlayers - 1).Cells(i, 1).AddressLocal, yhatNamesRange.Cells(j, 1).AddressLocal, ws, W.Cells(i, j).AddressLocal
      End If
    Next j
  Next i
End Sub

Function isContainedBy(rngSmall As Range, rngBig As Range) As Boolean
  isContainedBy = False
  If Intersect(rngSmall, rngBig) Is Nothing Then Exit Function
  If Intersect(rngSmall, rngBig).AddressLocal = rngSmall.AddressLocal Then isContainedBy = True
  
End Function

Sub initRpropNextWeights(ws As Worksheet)
  Dim i&, j&
  For i = 1 To ws.Range("Weights").Rows.Count
    For j = 1 To ws.Range("Weights").Columns.Count
      If Trim(ws.Range("Weights").Cells(i, j).Formula) <> "" Then
        If ws.Range("Weights").Cells(i, j).hasFormula Then
          ws.Range("rpropNextWeights").Cells(i, j).Formula = ws.Range("Weights").Cells(i, j).Formula
        Else
          '=IF(method = ""rprop-"", $L$3 - SIGN($B$3) * $L$64, IF(method= ""bp"", $L$3 - $B$3 * learningRate, IF(method=""rprop"", $L$3 - SIGN($B$3) * $L$64, NA())))
          ws.Range("rpropNextWeights").Cells(i, j).Formula = Replace(Replace(Replace( _
            "=IF(method = ""rprop-"", $L$3 - SIGN($B$3) * $L$64, IF(method= ""bp"", $L$3 - $B$3 * learningRate, IF(method=""rprop"", $L$3 - SIGN($B$3) * $L$64, NA())))", _
            "$L$3", ws.Range("Weights").Cells(i, j).AddressLocal), _
            "$B$3", ws.Range("Grads").Cells(i, j).AddressLocal), _
            "$L$64", ws.Range("rprop").Cells(i, j).AddressLocal)
        End If
      End If
    Next j
  Next i
End Sub
