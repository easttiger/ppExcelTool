VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtDoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
  Dim i&, j&, k&, n&, layers, nrow&, ans
  n = [A2].End(xlDown).Value
  
  For i = 1 To n
    shtPrep.Range("HiddenLayers").Formula = ""
    [A1] = i
    Me.Calculate
    shtPrep.Calculate
    ans = Split([C1], ":")
    ReDim layers(LBound(ans) To UBound(ans)) As Long
    For k = LBound(layers) To UBound(layers)
      layers(k) = CLng(ans(k))
    Next k
    
    nrow = Application.Max(layers) + 2
    For k = 1 To nrow - shtPrep.Range("HiddenLayers").Rows.Count
      shtPrep.cbtnHiddenLayersAddRow_Click
    Next k
    For k = 1 To UBound(layers) - shtPrep.Range("HiddenLayers").Columns.Count
      shtPrep.cbtnHiddenLayersAddCol_Click
    Next k
    shtPrep.Range("HiddenLayers").Cells(1, 1).Formula = "=nn(""1"")"
    shtPrep.Range("HiddenLayers").Cells(2, 1).Resize(layers(0), 1).Formula = "=nn(""logit"")"
    If UBound(layers) > 1 Then
      For j = 2 To UBound(layers)
        shtPrep.Range("HiddenLayers").Cells(1, j).Formula = "=nn(""1"")"
        shtPrep.Range("HiddenLayers").Cells(2, j).Resize(layers(j - 1), 1).Formula = "=nn(""logit"")"
      Next j
    End If
    Call shtPrep.cbtn_gen_Click
    With ActiveSheet
      .Range("epoch").Value = 500
      .Range("method").Value = Me.[G1]
      .Range("learningRate").Value = Me.[J1]
      .Range("batch_size").Value = Me.[H1]
      .Range("DO_TRACKING").Value = "loss,tloss,Grads,Weights"
      Call train(ActiveSheet)
      .Visible = xlHidden
    End With
    Call txt_write(CStr(i), path_dropbox() & "progress_" & ThisWorkbook.Name & ".txt")
    ThisWorkbook.Save
  Next i
End Sub


Private Sub CommandButton2_Click()
  Dim r As Range, c As Range, i&
  shtTrace.Activate
  For i = 1 To Me.Range("A2").End(xlDown).Value
    Set r = Me.Range("K2").Offset(i, 0)
    Call shtTrace.makeChart(r.Value, "Training Error " & i, 500 * ((i - 1) \ 8) + 10, 500 * ((i - 1) Mod 8) * 2 + 10)
    Call shtTrace.makeChart(r.Offset(0, 1).Value, "Testing Error " & i, 500 * ((i - 1) \ 8) + 10, 500 * (((i - 1) Mod 8) * 2 + 1) + 10)
    Set r = r.Offset(1, 0)
  Next i
End Sub
