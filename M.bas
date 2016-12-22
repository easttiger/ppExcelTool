Attribute VB_Name = "M"
Option Explicit

Function rnorm1(Optional mu As Double = 0#, Optional sd As Double = 1#)
  rnorm1 = Application.Evaluate("=NORM.INV(RAND()," & mu & "," & sd & ")")
End Function

Function xlActfuncFormula(ByVal strFuncName As String, Optional ByVal arg1 As String, _
                    Optional ByVal arg2 As String, Optional ByVal arg3 As String)
  Dim res As String
  Select Case LCase(Trim(strFuncName))
    Case "logistic", "logit":
      'logit = tanh(x / 2) / 2 + 0.5
      res = "=TANH(MMULT(TRANSPOSE(" & arg1 & ")," & arg2 & ")/2)/2+0.5"
    Case "mlogit", "o:mlogit":
      res = "=EXP(" & arg1 & ")/MMULT(TRANSPOSE(" & arg2 & "),EXP(" & arg1 & "))"
    Case "lin", "linear", "o:lin", "o:linear":
      res = "=MMULT(TRANSPOSE(" & arg2 & ")," & arg1 & ")"
  End Select
  xlActfuncFormula = res
End Function

Function xlWeiGradFormula(ByVal strLossFunc As String, ByVal totNumLayers As Byte, ByVal layerLevel As Byte)
  Dim res As String, i&, j&, k&
lbl_xentropy:
  If strLossFunc = "xen" And layerLevel <= totNumLayers And layerLevel >= 1 Then
    res = "D_" & totNumLayers & "*(1-D_" & totNumLayers & ")*(yobs-yhat)"
    For k = totNumLayers - 1 To layerLevel Step -1
      res = "D_" & k & "*(1-D_" & k & ")*MMULT(W_" & k + 1 & "," & res & ")"
    Next k
    res = "=-MMULT(D_" & layerLevel - 1 & ",TRANSPOSE(" & res & "))"
    xlWeiGradFormula = res: Exit Function
  End If
  
lbl_L2:
  If strLossFunc = "L2" And layerLevel <= totNumLayers + 1 And layerLevel >= 1 Then
    res = "(yhat-yobs)*2"
    For k = totNumLayers To layerLevel Step -1
      res = "D_" & k & "*(1-D_" & k & ")*MMULT(W_" & (k + 1) & "," & res & ")"
    Next k
    res = "=MMULT(D_" & layerLevel - 1 & ",TRANSPOSE(" & res & "))"
    xlWeiGradFormula = res: Exit Function
  End If
  
e:
  xlWeiGradFormula = "=NA()"
End Function

Sub train(ws As Worksheet)
On Error GoTo 0
  Dim timer As Double: timer = Now()
  Dim lr As Double: lr = ws.Range("learningRate").Value
  Dim szTrain As Long: szTrain = ws.Range("D_0i").Columns.Count
  Dim szBatch As Long: szBatch = szTrain
  If Application.IsNumber(ws.Range("batch_size").Value) Then
    szBatch = ws.Range("batch_size").Value + 0
    If szBatch > szTrain Then szBatch = szTrain
  End If
  
  Dim repBatch As Long: repBatch = ws.Range("batch_steps").Value
  Dim epoch As Long: epoch = ws.Range("epoch").Value
  Dim d As Long: d = ws.Range("Grads").Cells(1, 1).Column - ws.Range("Weights").Cells(1, 1).Column
  Dim u As Long: u = ws.Range("prevState").Cells(1, 1).Row - ws.Range("Weights").Cells(1, 1).Row
  Dim r As Range, i As Long, j As Long, loss_last As Double, loss_start As Double, loss_lastEpoch_train As Double, loss_lastEpoch_test As Double
  Dim nLayers As Long: nLayers = ws.Range("nLayers").Value
  Dim roll As Long: roll = ws.Range("roll").Value
lbl_vbmemory:
  For Each r In ws.Range("Weights").Cells
    If r.hasFormula Then j = j + 1
  Next r
  If j > 0 Then
  Dim fmlacells: ReDim fmlacells(1 To j, 1 To 3)
    For Each r In ws.Range("Weights").Cells
      If r.hasFormula Then
        fmlacells(j, 1) = r.Row - ws.Range("Weights").Cells(1, 1).Row + 1
        fmlacells(j, 2) = r.Column - ws.Range("Weights").Cells(1, 1).Column + 1
        fmlacells(j, 3) = r.FormulaLocal
        j = j - 1
      End If
    Next r
  End If
lbl_archive:
  loss_start = ws.Range("totloss").Value
  'loss_best = loss_start
lbl_run_new:
  Call initNextWeights(ws)

  Dim nrowsBatch As Long, iRoll As Long, nRoll As Long, iRepBatch As Long
  nRoll = Int((szTrain - szBatch) / roll)
  ws.Calculate
  While epoch > 0
    loss_lastEpoch_train = ws.Range("totloss").Value
    loss_lastEpoch_test = ws.Range("totloss_t").Value
lbl_RMS_reset_initial_accumulator:
    If ws.Range("method").Value Like "rmsprop*" Then
      For Each r In ws.Range("prevRMSPROP").Cells
        If IsNumeric(r.Value) And Not IsEmpty(r.Formula) Then
          r.Value = 0
        End If
      Next r
    End If
lbl_SGD_init_batch:
    For j = 0 To nLayers
      nrowsBatch = ws.Range("D_" & j & "i").Rows.Count
      ws.Names("D_" & j).RefersTo = "='" & ws.Name & "'!" & ws.Range("D_" & j & "i").Cells(1, 1).resize(nrowsBatch, szBatch).AddressLocal
    Next j
    nrowsBatch = ws.Range("yhati").Rows.Count
    ws.Names("yhat").RefersTo = "='" & ws.Name & "'!" & ws.Range("yhati").Cells(1, 1).resize(nrowsBatch, szBatch).AddressLocal
    nrowsBatch = ws.Range("yobsi").Rows.Count
    ws.Names("yobs").RefersTo = "='" & ws.Name & "'!" & ws.Range("yobsi").Cells(1, 1).resize(nrowsBatch, szBatch).AddressLocal
    ws.Names("loss").RefersTo = "='" & ws.Name & "'!" & ws.Range("lossi").Cells(1, 1).resize(1, szBatch).AddressLocal
    For iRoll = 0 To nRoll Step 1
      loss_last = ws.Range("totloss").Value
      For iRepBatch = 1 To repBatch
lbl_update:
        ws.Calculate
        Application.StatusBar = "Epoch:" & epoch & " Batch:" & iRoll & " Step:" & iRepBatch & " |Previous epoch's exit losses: test=" & loss_lastEpoch_test & " train=" & loss_lastEpoch_train
'        If ws.Range("totloss") < loss_best Then
'          loss_best = ws.Range("totloss")
'          ws.Range("bestWeights").Value2 = ws.Range("Weights").Value2
'        End If
        ws.Range("prevState").Value2 = ws.Range("WorkRange").Value2
        ws.Range("Weights").Value2 = ws.Range("nextWeights").Value2
        For j = LBound(fmlacells, 1) To UBound(fmlacells, 1)
          ws.Range("Weights").Cells(fmlacells(j, 1), fmlacells(j, 2)).FormulaLocal = fmlacells(j, 3)
        Next j
lbl_post_update:
        Select Case ws.Range("method").Value
          Case "bp": 'nothing to do
          
          Case "rprop-": 'no weight backtracking
            
            ws.Range("prevRPROP").Value2 = ws.Range("rprop").Value2
            
          Case "rprop": 'has weight backtracking
            ws.Calculate
            If ws.Range("totloss") >= loss_last Then
              
              For Each r In ws.Range("Weights").Cells
                r.Select
                If Trim(r.Formula) <> "" And IsNumeric(r.Formula) Then
                  If Sgn(r.Offset(0, d).Value) <> Sgn(r.Offset(u, d).Value) Then
                    r.Value = r.Offset(u, 0).Value
                  End If
                End If
              Next r
            End If
            ws.Range("prevRPROP").Value2 = ws.Range("rprop").Value2
            
          Case "rmsprop": 'decaying history
            ws.Calculate
'            If ws.Range("totloss") >= loss_last Then
'              For Each r In ws.Range("Weights").Cells
'                If Trim(r.Formula) <> "" And IsNumeric(r.Formula) Then
'                  If Sgn(r.Offset(0, d).Value) <> Sgn(r.Offset(u, d).Value) Then
'                    r.Value = r.Offset(u, 0).Value
'                  End If
'                End If
'              Next r
'            End If
            ws.Range("prevRMSPROP").Value2 = ws.Range("rmsprop").Value2
          Case Else:
          
        End Select
      Next iRepBatch
lbl_SGD_roll_batch:
      
      nrowsBatch = ws.Range("D_0i").Rows.Count
      ws.Names("D_0").RefersTo = "='" & ws.Name & "'!" & ws.Range("D_0").Offset(0, roll).AddressLocal
      
      nrowsBatch = ws.Range("yobsi").Rows.Count
      ws.Names("yobs").RefersTo = "='" & ws.Name & "'!" & ws.Range("yobs").Offset(0, roll).AddressLocal
      
    Next iRoll
'    If ws.Range("useBestWeightsPerEpoch") = True Then
'      ws.Range("Weights").Value2 = ws.Range("bestWeights").Value2
'      For j = LBound(fmlacells, 1) To UBound(fmlacells, 1)
'          ws.Range("Weights").Cells(fmlacells(j, 1), fmlacells(j, 2)).FormulaLocal = fmlacells(j, 3)
'        Next j
'      ws.Calculate
'    End If
lbl_SGD_restore_fullbatch:
    For j = 0 To nLayers
      nrowsBatch = ws.Range("D_" & j & "i").Rows.Count
      ws.Names("D_" & j).RefersTo = "='" & ws.Name & "'!" & ws.Range("D_" & j & "i").AddressLocal
    Next j
    nrowsBatch = ws.Range("yhati").Rows.Count
    ws.Names("yhat").RefersTo = "='" & ws.Name & "'!" & ws.Range("yhati").AddressLocal
    nrowsBatch = ws.Range("yobsi").Rows.Count
    ws.Names("yobs").RefersTo = "='" & ws.Name & "'!" & ws.Range("yobsi").AddressLocal
    ws.Names("loss").RefersTo = "='" & ws.Name & "'!" & ws.Range("lossi").AddressLocal
    
    ws.Calculate
    epoch = epoch - 1
  Wend
'  For j = LBound(fmlacells, 1) To UBound(fmlacells, 1)
'    ws.Range("Weights").Cells(fmlacells(j, 1), fmlacells(j, 2)).FormulaLocal = fmlacells(j, 3)
'  Next j
  
lbl_msgbox:
  Dim msg As String
  msg = "Trained " & ws.Range("epoch").Value & " epochs of "
  ws.Calculate
  Select Case ws.Range("method").Value
    Case "bp:"
      msg = "Backprop with learning rate " & lr
    Case "rprop-":
      msg = "rprop- (without weight backtracking) with learning rate resilience scalars {" & ws.Range("rpropdn").Value & ", " & ws.Range("rpropup").Value & _
            "} and learning rate bounds [" & ws.Range("rpropfloor").Value & " to " & ws.Range("rpropcap").Value & "]"
    Case "rprop":
      msg = "rprop with learning rate resilience scalars {" & ws.Range("rpropdn").Value & ", " & ws.Range("rpropup").Value & _
            "} and learning rate bounds [" & ws.Range("rpropfloor").Value & " to " & ws.Range("rpropcap").Value & "]"
  End Select
  MsgBox "Trained " & ws.Range("epoch").Value & " epochs of " & msg & vbCr & "Time spent:  " & Format(Now() - timer, "hh:mm:ss") & " (hh:mm:ss)" & vbCr & _
        "loss before = " & loss_start & vbCr & "loss now = " & ws.Range("totloss").Value
End Sub

