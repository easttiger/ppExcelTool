VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtPrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cbtn_clearShapes_Click()
  clearShapes Me
End Sub

Private Sub cbtn_gen_Click()
On Error GoTo e
  Dim xlcalcb4: xlcalcb4 = Application.Calculation
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Dim ans, x, i&, j&, k&, h&, n&, ntest&, M&, r As Range, c As Range, q As Range
lbl_initialize:
  ReDim x(1 To Sheets.Count)
  For i = 1 To Sheets.Count
    x(i) = Sheets(i).Name
  Next i
  Sheet0.Copy after:=Me
  Dim shtDes As Worksheet
  For i = 1 To Sheets.Count
    For j = 1 To UBound(x)
      If x(j) = Sheets(i).Name Then GoTo lbl_next_i
    Next j
    Set shtDes = Sheets(i)
    Exit For
lbl_next_i:
  Next i
    
  With shtDes.Cells
      .Interior.Color = RGB(0, 0, 0)
      .Font.ThemeColor = xlThemeColorDark1
      .Font.TintAndShade = 0
      .NumberFormat = "0.###"
  End With
lbl_get_hidden_layer_width:
  Dim hcol As Range, h_dim(1 To 20) As Long
  j = 0
  For Each hcol In Me.Range("HiddenLayers").Columns
    ans = Application.CountA(hcol)
    j = j + 1: h_dim(3 + j) = ans
    h_dim(3) = h_dim(3) + ans 'stores total hidden neuron count
    If ans > 0 Then h_dim(2) = h_dim(2) + 1 'stores num hidden layers
    h_dim(1) = Application.Max(h_dim(1), ans) 'stores max hidden layer width
  Next hcol
lbl_final_transform:
  Set hcol = Me.Range("final_transform")
  ans = Application.CountA(hcol)
  j = h_dim(2) + 1: h_dim(3 + j) = ans
  h_dim(3) = h_dim(3) + ans 'stores total hidden neuron count
  If ans > 0 Then h_dim(2) = h_dim(2) + 1 'stores num hidden layers
  h_dim(1) = Application.Max(h_dim(1), ans) 'stores max hidden layer width
lbl_get_xycount:
  Dim nx As Byte: nx = Application.Evaluate("=COUNTA(" & shtPrep.Name & "!InputData)")
  Dim ny As Byte: ny = Application.Evaluate("=COUNTA(" & shtPrep.Name & "!yhat)")
lbl_set_road:
  Dim road As Range: Set road = shtDes.Columns(2 + _
                                h_dim(1) * 2).Cells(3, 1).resize(h_dim(2) + ny * 2 + 1, 1)
lbl_set_buttons:
  Dim btnRange As Range: Set btnRange = road.Cells(1, 1).Offset(-2, 0)
  With btnRange.resize(2, 1)
    .Merge
    .EntireRow.HorizontalAlignment = xlCenter
    .EntireRow.VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
    .Cells(1, 1).Value = "Train"
    .Cells(1, 1).Font.Italic = True
    .Cells(1, 1).Name = "'" & shtDes.Name & "'!train"
    Call colorRangeBorder(.Cells, 10)
    
    .Cells(1, 2).Value = "Repeat"
    With .Cells(2, 2)
      .Value = 100
      .Name = "'" & shtDes.Name & "'!repeat"
      .Font.Color = RGB(255, 0, 0)
      .NumberFormat = "General"
    End With
    
    .Cells(1, 3).Value = "Method"
    With .Cells(2, 3)
      .Value = "rprop-"
      .Name = "'" & shtDes.Name & "'!method"
      .Font.Color = RGB(255, 0, 0)
      With .Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="bp,rprop,rprop-"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
      End With
    End With
    
    .Cells(1, 4).Value = "lrate"
    With .Cells(2, 4)
      .Value = 0.01
      .Name = "'" & shtDes.Name & "'!learningRate"
      .Font.Color = RGB(255, 0, 0)
    End With
    
    .Cells(1, 5).Value = "rprop up"
    With .Cells(2, 5)
      .Value = 1.2
      .Name = "'" & shtDes.Name & "'!rpropup"
      .Font.Color = RGB(255, 0, 0)
    End With
    .Cells(1, 6).Value = "rprop dn"
    With .Cells(2, 6)
      .Value = 0.5
      .Name = "'" & shtDes.Name & "'!rpropdn"
      .Font.Color = RGB(255, 0, 0)
    End With
    .Cells(1, 7).Value = "rprop cap"
    With .Cells(2, 7)
      .Value = 50
      .Name = "'" & shtDes.Name & "'!rpropcap"
      .Font.Color = RGB(255, 0, 0)
    End With
    .Cells(1, 8).Value = "rprop floor"
    With .Cells(2, 8)
      .Value = 0.000001
      .Name = "'" & shtDes.Name & "'!rpropfloor"
      .Font.Color = RGB(255, 0, 0)
      .NumberFormat = "General"
    End With
      
  End With
  With btnRange.Offset(0, -1).resize(2, 1)
    .Merge
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
    .Cells(1, 1).Value = "Noisify"
    .Cells(1, 1).Font.Italic = True
    .Cells(1, 1).Name = "'" & shtDes.Name & "'!imputeNoise"
    Call colorRangeBorder(.Cells, 10)
  End With
  With btnRange.Offset(0, -2).resize(2, 1)
    .Merge
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
    .Cells(1, 1).Value = "Trim"
    .Cells(1, 1).Font.Italic = True
    .Cells(1, 1).Name = "'" & shtDes.Name & "'!trimWeiMats"
    Call colorRangeBorder(.Cells, 10)
  End With
  With btnRange.Offset(0, -3).resize(2, 1)
    .Merge
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
    .Cells(1, 1).Value = "Remove Plot"
    .Cells(1, 1).Font.Italic = True
    .Cells(1, 1).Name = "'" & shtDes.Name & "'!removePlot"
    Call colorRangeBorder(.Cells, 10)
  End With
  With btnRange.Offset(0, -4).resize(2, 1)
    .Merge
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
    .Cells(1, 1).Value = "Plot Net"
    .Cells(1, 1).Font.Italic = True
    .Cells(1, 1).Name = "'" & shtDes.Name & "'!plotNet"
    Call colorRangeBorder(.Cells, 10)
  End With
  
  With btnRange.Offset(0, -5).resize(2, 1)
    .Cells(1, 1).Value = "#Weights"
    .Cells(1, 1).Font.Italic = False
    .Cells(2, 1).Name = "'" & shtDes.Name & "'!nWeights"
    .Cells(2, 1).FormulaLocal = "=COUNTA(Weights)-COUNTIF(Weights,0)"
  End With
  
  With btnRange.Offset(0, -6).resize(2, 1)
    .Merge
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
    .Cells(1, 1).Value = "Shrink"
    .Cells(1, 1).Font.Italic = True
    .Cells(1, 1).Name = "'" & shtDes.Name & "'!shrinkWeiMats"
    Call colorRangeBorder(.Cells, 10)
  End With

lbl_determine_loss_function:
  Dim strLossFunction As String
  Select Case Trim(Me.Range("loss").Value)
    Case "e:L2": strLossFunction = "L2"
    Case "e:xentropy": strLossFunction = "xen"
    Case Else: Err.Raise 1, , "Unknown Loss Function"
  End Select
lbl_initialize_k:
  k = 0
  n = Me.Range("ntrain").Value
  ntest = Me.Range("ntest").Value
lbl_parse_ii:
  Dim dr As Range, pr As Range
  For Each r In Me.Range("InputData").Cells
    If Trim(r.FormulaLocal) <> "" Then
      k = k + 1
      Set c = road.Cells(k, 1): c.Value = r.Value
      If CStr(r.Value) = "1" Then
        c.Offset(0, 1).resize(1, n + ntest).Value = 1
      ElseIf r.FormulaLocal Like "=ii(*)" Then
        Set pr = pre(r)
        c.Offset(0, 1).resize(1, pr.Rows.Count).FormulaArray = "=TRANSPOSE(" & shtPrep.Name & "!" & pr.AddressLocal & ")"
        c.Offset(0, 1).resize(1, pr.Rows.Count).Value2 = c.Offset(0, 1).resize(1, pr.Rows.Count).Value2
      Else
      
      End If
    End If
  Next r
  road.Cells(1, 1).resize(nx, 1).Name = "'" & shtDes.Name & "'!N_0"
  road.Cells(1, 1).Offset(0, 1).resize(nx, n).Name = "'" & shtDes.Name & "'!D_0"
  road.Cells(1, 1).Offset(0, 1 + n).resize(nx, ntest).Name = "'" & shtDes.Name & "'!E_0"
  Dim v As Range: Set v = shtDes.Range("D_0")
  Call colorRangeBorder(v, 9)
  Dim u As Range: Set u = shtDes.Range("E_0") 'test
  Call colorRangeBorder(u, 8)
lbl_initialize_W_matrix:
  Dim W As Range, Wpre As Range, DW As Range
  Set W = road.Cells(1, 1).Offset(0, -h_dim(3 + 1)).resize(nx, h_dim(3 + 1))
  For i = 1 To W.Rows.Count
    W.Rows(i).Formula = "=RNORM1()/ AVERAGE(" & shtDes.Range("D_0").Rows(i).AddressLocal & ")"
  Next i
  W.Value2 = W.Value2 'make const
  W.Name = "'" & shtDes.Name & "'!W_1"
  'Set DW = Intersect(shtDes.Columns(1), W.EntireRow).Cells(1, 1).Resize(W.Rows.Count, W.Columns.Count)
  Set DW = W.Offset(0, -1 - h_dim(1))
  DW.Name = "'" & shtDes.Name & "'!G_1"
  Call colorRangeBorder(W, 6)
  Call colorRangeBorder(DW, 8)
  Call underscoreFormulaCells(W)
lbl_set_1st_col:
'  W.Columns(1).Value = 0:  W.Cells(1, 1).Value = 999  'need to change if non-sigmoidal
  
lbl_parse_hiddenlayers:
  j = 0
  For Each hcol In Me.Range(Me.Range("HiddenLayers"), Me.Range("final_transform")).Columns
    j = j + 1
    If j <= h_dim(2) Then
      If j > Me.Range("HiddenLayers").Count Then Debug.Assert hcol.AddressLocal = shtPrep.Range("final_transform").AddressLocal
      road.Cells(k, 1).Offset(1, 0).resize(h_dim(3 + j), 1).Name = "'" & shtDes.Name & "'!N_" & j
      ans = v.Name.Name
      Set v = road.Cells(k, 1).Offset(1, 1).resize(h_dim(3 + j), n)
      v.FormulaArray = xlActfuncFormula("logit", W.Name.Name, ans)
      v.Name = "'" & shtDes.Name & "'!D_" & j
      Call colorRangeBorder(v, 10)
      'test data
      ans = u.Name.Name
      Set u = road.Cells(k, 1).Offset(1, 1 + n).resize(h_dim(3 + j), ntest)
      u.FormulaArray = xlActfuncFormula("logit", W.Name.Name, ans)
      u.Name = "'" & shtDes.Name & "'!E_" & j
      Call colorRangeBorder(u, 8)
      
      Set Wpre = W
      If j < h_dim(2) Then
        Set W = road.Cells(k, 1).Offset(1, -h_dim(3 + j + 1)).resize(h_dim(3 + j), h_dim(3 + j + 1))
        W.Formula = "=RNORM1(0,0.1)"
      Else
        'to output
        Set W = road.Cells(k, 1).Offset(1, -ny).resize(h_dim(3 + j), ny)
        W.Formula = "=RNORM1(0,0.1)"
      End If
      W.Value2 = W.Value2 'make const"
      W.Name = "'" & shtDes.Name & "'!W_" & (j + 1)
      'Set DW = Intersect(shtDes.Columns(1), W.EntireRow).Cells(1, 1).Resize(W.Rows.Count, W.Columns.Count)
      Set DW = W.Offset(0, -1 - h_dim(1))
      DW.Name = "'" & shtDes.Name & "'!G_" & (j + 1)
      Call colorRangeBorder(W, 6)
      Call colorRangeBorder(DW, 8)
      Call underscoreFormulaCells(W)
      
    End If
    For Each r In hcol.Cells
      If Trim(r.FormulaLocal) <> "" Then
        k = k + 1
        Set c = road.Cells(k, 1): c.Value = r.Value
        If CStr(c.Value) = "1" Then
          Wpre.Columns(1).FormulaLocal = "=0"
          Wpre.Cells(1, 1).FormulaLocal = "=999"
        End If
      End If
    Next r
    
  Next hcol
lbl_parse_yhat:
  Set q = road.Cells(k + 1, 1)
  q.Offset(0, -1).Value = "yhat"
  q.Offset(0, 1).resize(ny, n).Name = "'" & shtDes.Name & "'!yhat"
  Call colorRangeBorder(shtDes.Range("yhat"), 5)
  q.Offset(0, 1 + n).resize(ny, ntest).Name = "'" & shtDes.Name & "'!that"
  Call colorRangeBorder(shtDes.Range("that"), 8)
  For Each r In Me.Range("yhat").Cells
    If Trim(r.FormulaLocal) <> "" Then
      k = k + 1
      Set c = road.Cells(k, 1): c.Value = r.Value
    End If
  Next r
  
  If q.Value = "o:mlogit" Then
    W.FormulaLocal = "=1"
    Call colorRangeBorder(W, 5)
  End If
  q.Offset(0, 1).resize(ny, n).FormulaArray = xlActfuncFormula(q.Value, v.Name.Name, W.AddressLocal)
  q.Offset(0, 1 + n).resize(ny, ntest).FormulaArray = xlActfuncFormula(q.Value, u.Name.Name, W.AddressLocal)
lbl_parse_lossfunction:
  Set r = Me.Range("loss")

  If Trim(r.FormulaLocal) Like "=ee(*)" Then
   
    k = k + 1
    Set c = road.Cells(k, 1)
    c.Font.Color = RGB(255, 0, 0)
    c.Offset(0, -1).Value = "loss:train"
    c.Name = "'" & shtDes.Name & "'!totloss"
    c.Offset(0, 1).resize(1, n).Name = "'" & shtDes.Name & "'!loss"
    c.Offset(0, -2).Font.Color = RGB(255, 255, 0)
    c.Offset(0, -3).Value = "loss:test"
    c.Offset(0, -2).Name = "'" & shtDes.Name & "'!totloss_t"
    c.Offset(0, 1 + n).resize(1, ntest).Name = "'" & shtDes.Name & "'!tloss"
  Else
    Err.Raise 1, , "no loss function defined"
  End If
lbl_parse_yobs:
  Set q = road.Cells(k + 1, 1)
  q.Offset(0, -1).Value = "yobs"
  q.Offset(0, 1).resize(ny, n).Name = "'" & shtDes.Name & "'!yobs"
  Call colorRangeBorder(shtDes.Range("yobs"), 9)
  q.Offset(0, 1 + n).resize(ny, ntest).Name = "'" & shtDes.Name & "'!tobs"
  Call colorRangeBorder(shtDes.Range("tobs"), 8)
  For Each r In Me.Range("yobs").Cells
    If Trim(r.FormulaLocal) <> "" Then
      k = k + 1
      Set c = road.Cells(k, 1): c.Value = r.Value
      If r.FormulaLocal Like "=yy(*)" Then
        Set pr = pre(r)
        c.Offset(0, 1).resize(1, pr.Rows.Count).FormulaArray = "=TRANSPOSE(" & shtPrep.Name & "!" & pr.AddressLocal & ")"
        c.Offset(0, 1).resize(1, pr.Rows.Count).Value2 = c.Offset(0, 1).resize(1, pr.Rows.Count).Value2
      End If
    End If
  Next r
lbl_fill_observationwise_loss:
  For Each r In shtDes.Range("loss").Cells
    Select Case strLossFunction
      Case "xen":
        r.FormulaLocal = "=-SUMPRODUCT(LN(" & Intersect(r.EntireColumn, shtDes.Range("yhat")).AddressLocal & ")," & Intersect(r.EntireColumn, shtDes.Range("yobs")).AddressLocal & ")"
      Case "L2":
        r.FormulaArray = "=SUMSQ(" & Intersect(r.EntireColumn, shtDes.Range("yhat")).AddressLocal & " - " & Intersect(r.EntireColumn, shtDes.Range("yobs")).AddressLocal & ")"
    End Select
  Next r
  
  For Each r In shtDes.Range("tloss").Cells
    Select Case strLossFunction
      Case "xen":
        r.FormulaLocal = "=-SUMPRODUCT(LN(" & Intersect(r.EntireColumn, shtDes.Range("that")).AddressLocal & ")," & Intersect(r.EntireColumn, shtDes.Range("tobs")).AddressLocal & ")"
      Case "L2":
        r.FormulaArray = "=SUMSQ(" & Intersect(r.EntireColumn, shtDes.Range("that")).AddressLocal & " - " & Intersect(r.EntireColumn, shtDes.Range("tobs")).AddressLocal & ")"
    End Select
  Next r
lbl_fill_totloss:
  shtDes.Range("totloss").FormulaLocal = "=SUM(loss)/COUNT(loss)"
  shtDes.Range("totloss_t").FormulaLocal = "=SUM(tloss)/COUNT(tloss)"
lbl_fill_gradients:
  
  For j = h_dim(2) + 1 To 1 Step -1
    shtDes.Range("G_" & j).FormulaArray = xlWeiGradFormula(strLossFunction, h_dim(2), j)
  Next j
lbl_name_big_regions:
  shtDes.Range(shtDes.Range("W_1"), shtDes.Range("W_" & h_dim(2) + 1)).Name = "'" & shtDes.Name & "'!Weights"
  shtDes.Range(shtDes.Range("G_1"), shtDes.Range("G_" & h_dim(2) + 1)).Name = "'" & shtDes.Name & "'!Grads"
  shtDes.Range(shtDes.Range("G_1"), shtDes.Range("W_" & h_dim(2) + 1)).Name = "'" & shtDes.Name & "'!WorkRange"
  
  
  'previous state snapshot
  With shtDes.UsedRange(shtDes.UsedRange.Count).Offset(2, 0).End(xlToLeft)
    .Value = "Previous"
    .Offset(1, 0).resize(shtDes.Range("WorkRange").Rows.Count, shtDes.Range("WorkRange").Columns.Count).Name = "'" & shtDes.Name & "'!prevState"
  End With
  shtDes.Range("WorkRange").Copy: shtDes.Range("prevState").PasteSpecial xlPasteFormats
  'rprop weights back tracking
  k = shtDes.Range("W_1").Cells(1, 1).Column - shtDes.Range("G_1").Cells(1, 1).Column
  
  With shtDes.UsedRange(shtDes.UsedRange.Count).Offset(2, 0).End(xlToLeft)
    .Value = "rprop change size : prev"
    .Offset(0, k).Value = "rprop change size : current"
    .Offset(1, 0).resize(shtDes.Range("WorkRange").Rows.Count, shtDes.Range("WorkRange").Columns.Count).Name = "'" & shtDes.Name & "'!rpropUse"
  End With
  shtDes.Range("WorkRange").Copy: shtDes.Range("rpropUse").PasteSpecial xlPasteFormats
  shtDes.Range("WorkRange").Copy: shtDes.Range("rpropUse").PasteSpecial xlPasteValues
  shtDes.Range("rpropUse").Cells(1, 1).resize(shtDes.Range("Grads").Rows.Count, shtDes.Range("Grads").Columns.Count).Name = "'" & shtDes.Name & "'!prevRPROP"
  shtDes.Range("rpropUse").Cells(1, 1 + k).resize(shtDes.Range("Weights").Rows.Count, shtDes.Range("Weights").Columns.Count).Name = "'" & shtDes.Name & "'!rprop"
  h = shtDes.Range("rpropUse").Rows.Count + 2
  shtDes.Range("rpropUse").Cells(h, 1 + k).resize(shtDes.Range("Weights").Rows.Count, shtDes.Range("Weights").Columns.Count).Name = "'" & shtDes.Name & "'!rpropNextWeights"
  shtDes.Range("Weights").Copy: shtDes.Range("rpropNextWeights").PasteSpecial xlPasteFormats
  
  Call initRpropNextWeights(shtDes)
  
  Dim grad2prevRPROP As Long: grad2prevRPROP = shtDes.Range("Grads").Cells(1, 1).Row - shtDes.Range("prevRPROP").Cells(1, 1).Row
  Dim prevGrad2prevRPROP As Long: prevGrad2prevRPROP = shtDes.Range("prevState").Cells(1, 1).Row - shtDes.Range("prevRPROP").Cells(1, 1).Row
  For Each r In shtDes.Range("prevRPROP").Cells
    If IsNumeric(r.Value) And Not Trim(r.FormulaLocal) = "" Then
      '=MIN(rpropcap,MAX(rpropfloor, IF(SIGN(A45)*SIGN(A3)>0,rpropup,IF(SIGN(A45)*SIGN(A3)<0,rpropdn,0))*A81))
      r.Offset(0, k).FormulaLocal = Replace(Replace(Replace("=MIN(rpropcap,MAX(rpropfloor, IF(SIGN(A45)*SIGN(A3)>0,rpropup,IF(SIGN(A45)*SIGN(A3)<0,rpropdn,0))*A81))", _
                                            "A45", r.Offset(prevGrad2prevRPROP, 0).AddressLocal), _
                                            "A3", r.Offset(grad2prevRPROP, 0).AddressLocal), _
                                            "A81", r.AddressLocal)
    End If
  Next r
  shtDes.Calculate
  shtDes.Visible = xlSheetVisible
  shtDes.Select
e:
  Application.Calculation = xlcalcb4
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

Private Sub cbtn_getNeuron_Click()
  createNeuron Rnd() * 500, Rnd() * 300, Replace(ActiveCell.AddressLocal, "$", ""), , Me
End Sub


Private Sub cbtn_getAllNeurons_Click()
  Application.EnableEvents = False
  Dim r As Range
  For Each r In shtPrep.Range("DesignArea").Cells
    If r.Formula Like "=nn(*)" Or r.Formula Like "=ii(*)" Or r.Formula Like "=oo(*)" Or r.Formula Like "=*ee*(*)" Then
      Call createNeuron(500 + r.Column * 70, r.Row * 35, IIf(r.Value = "1", "1", Replace(r.AddressLocal, "$", "")), Replace(r.AddressLocal, "$", ""), Me)
    End If
  Next r
  For Each r In shtPrep.Range("DesignArea").Cells
    If r.Formula Like "=nn(*)" Or r.Formula Like "=ii(*)" Or r.Formula Like "=oo(*)" Then
      Call emitConnector(r, Me)
    End If
  Next r
  
  Application.EnableEvents = True
End Sub


Private Sub cbtn_rerouteConnectors_Click()
    Application.EnableEvents = False
    
    Dim sh As Shape
    For Each sh In shtPrep.Shapes
        If sh.Name Like "*Connector*" Then
            sh.RerouteConnections
        End If
    Next sh
           
    Application.EnableEvents = True
End Sub


Private Sub cbtnHiddenLayersAddRow_Click()
  Dim r As Range:  Set r = shtPrep.Range("HiddenLayers").Item(shtPrep.Range("HiddenLayers").Cells.Count)
  r.EntireRow.Insert xlShiftDown, xlFormatFromLeftOrAbove
End Sub
Private Sub cbtnHiddenLayersAddCol_Click()
  Dim r As Range:  Set r = shtPrep.Range("final_transform")
  r.Insert xlShiftToRight, xlFormatFromLeftOrAbove
  shtPrep.Names("HiddenLayers").RefersToLocal = "=Prep!" & Range(shtPrep.Range("HiddenLayers"), shtPrep.Range("final_transform").Offset(0, -1)).AddressLocal
End Sub


Private Sub cbtnHiddenLayersFit_Click()
  Dim c As Range, r As Range, r2 As Range, nrow&, ncol&, n&
  nrow = 0
  For Each c In shtPrep.Range("DesignArea").Columns
      Set r = c.Cells(c.Rows.Count, 1).End(xlUp)
      n = r.Row - c.Cells(1, 1).Row + 1
      If nrow < n Then nrow = n
  Next c
  If nrow > 0 Then
    Set r = shtPrep.Range("DesignArea").Cells(1, 1).Offset(nrow, 0)
    Set r2 = shtPrep.Range("DesignArea")
    Set r2 = r2.Cells(r2.Rows.Count - 1, 1)
    If r2.Row >= r.Row Then shtPrep.Range(r, r2).EntireRow.Delete
  End If
  ncol = 0
  Set r2 = shtPrep.Range("final_transform").Offset(0, -1)
  Set c = r2
  While Application.CountA(c) = 0
    ncol = ncol + 1
    Set c = c.Offset(0, -1)
  Wend
  Set c = c.Offset(0, 1)
  If c.Column <= r2.Column Then
    shtPrep.Range(c, r2).Delete xlShiftToLeft
  End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
On Error GoTo e
  If Target.Cells(1, 1).hasFormula Then
    If Target.Cells(1, 1).FormulaLocal Like "=nn(*)" Or Target.Cells(1, 1).FormulaLocal Like "=ii(*)" Or Target.Cells(1, 1).FormulaLocal Like "=oo(*)" Or Target.Cells(1, 1).FormulaLocal Like "=*ee*(*)" Then
      Cancel = True
      Application.EnableEvents = False: Application.ScreenUpdating = False
      Shapes("Oval_" & Replace(Target.AddressLocal, "$", "")).Select
      Application.EnableEvents = True: Application.ScreenUpdating = True
      Exit Sub
    End If
  End If
  
  If Target.Cells(1, 1).AddressLocal = shtPrep.Range("paste_data_from_clipboard").Cells(1, 1).AddressLocal Then
    Cancel = True
    Call pasteDataFromClipboard
  End If
  
  If Target.Cells(1, 1).AddressLocal = shtPrep.Range("setup_input_output").Cells(1, 1).AddressLocal Then
    Cancel = True
    Call setup_input_output
  End If
e:
End Sub

Private Sub pasteDataFromClipboard()
  Application.EnableEvents = False
  Application.ScreenUpdating = False
On Error GoTo e
  shtPrep.Activate
  shtPrep.Range("dataStart").Offset(0, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Dim ncol As Long, nrow As Long
  ncol = Selection.Columns.Count
  nrow = Selection.Rows.Count
  Dim p(1 To 3) As Range
  Set p(1) = shtPrep.Range("dataStart").Offset(0, ncol + 1)
  Set p(2) = shtPrep.Range("dataStart").Offset(nrow, 0)
  Set p(3) = shtPrep.UsedRange(shtPrep.UsedRange.Count)
  If p(3).Column >= p(1).Column Then
    shtPrep.Range(p(1), p(3)).Value = ""
  End If
  If p(3).Row >= p(2).Row Then
    shtPrep.Range(p(2), p(3)).Value = ""
  End If
  shtPrep.Range("dataStart").Offset(1, 0).resize(nrow - 1, 1).Formula = "=IF(RAND() > 0.3, ""train"", ""test"")"
  shtPrep.Range("dataStart").Offset(1, 0).resize(nrow - 1, 1).Value2 = shtPrep.Range("dataStart").Offset(1, 0).resize(nrow - 1, 1).Value2
  shtPrep.Range("dataStart").Offset(-1, 1).resize(1, ncol).Value = ""
  shtPrep.ListObjects("tbl_data").resize shtPrep.Range("dataStart").resize(nrow, ncol + 1)
  With shtPrep.ListObjects("tbl_data").Sort
      .SortFields.Clear
      .SortFields.Add Key:=Range("tbl_data[[#All],[as]]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  shtPrep.Range(shtPrep.Range("start_in_out_spec"), shtPrep.Range("start_in_out_spec").Offset(0, 200).End(xlToLeft)).Value = ""
  shtPrep.Range("start_in_out_spec").Select
  Me.Calculate
e:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub

Private Sub setup_input_output()
  Dim r As Range, cin As Range, cout As Range
  Dim nmax As Long: nmax = Me.Range("InputData").Rows.Count
  Dim nin&, nout&
  nin = 1
  nout = 0
  
  For Each r In shtPrep.Range("tbl_data[#Headers]").Cells
    Select Case LCase(Trim(r.Offset(-1, 0).Value))
      Case "x": nin = nin + 1
      Case "y": nout = nout + 1
      Case Else:
    End Select
  Next r
  If nin > nmax Or nout > nmax Then
    MsgBox "Not enough rows in design region", vbExclamation
    Exit Sub
  End If
  
  
  shtPrep.Range("InputData").Value = ""
  Set cin = shtPrep.Range("InputData").Cells(1, 1)
  cin.FormulaLocal = "=nn(""1"")"
  shtPrep.Range("yobs").Value = ""
  Set cout = shtPrep.Range("yobs").Cells(1, 1)
  For Each r In shtPrep.Range("tbl_data[#Headers]").Cells
    Select Case LCase(Trim(r.Offset(-1, 0).Value))
      Case "x":
        Set cin = cin.Offset(1, 0)
        cin.FormulaLocal = Replace("=ii(tbl_data[[#Headers],[z1]],tbl_data[z1])", "z1", r.Value)
      Case "y":
        cout.FormulaLocal = Replace("=yy(tbl_data[[#Headers],[z7]],tbl_data[z7])", "z7", r.Value)
        Set cout = cout.Offset(1, 0)
      Case Else:
    End Select
  Next r
End Sub
