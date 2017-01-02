Attribute VB_Name = "M"
Option Explicit
Private Function EXIT_NOW() As Boolean
  EXIT_NOW = ActiveSheet.Range("EXIT_NOW")
End Function
Public Function DEBUG_LEVEL() As Integer
  DEBUG_LEVEL = ActiveSheet.Range("DEBUG_LEVEL")
End Function
Private Function DO_DROPOUT() As Boolean
  DO_DROPOUT = ActiveSheet.Range("DO_DROPOUT")
End Function
Function rnorm1(Optional mu As Double = 0#, Optional sd As Double = 1#)
  rnorm1 = Application.Evaluate("=NORM.INV(RAND()," & mu & "," & sd & ")")
End Function

Function xlActfuncFormula(ByVal strFuncName As String, Optional ByVal arg1 As String, _
                    Optional ByVal arg2 As String, Optional ByVal arg3 As String, _
                    Optional returnDerivativeFunction As Boolean = False)
  Dim res$
  Select Case LCase(Trim(strFuncName))
    Case "logistic", "logit":
      'logit = tanh(x / 2) / 2 + 0.5
      If returnDerivativeFunction Then
        res = arg1 & "*(1-" & arg1 & ")"
      Else
        res = "=TANH(MMULT(TRANSPOSE(" & arg1 & ")," & arg2 & ")/2)/2+0.5"
      End If
    Case "tanh":
      If returnDerivativeFunction Then
        res = "(1-" & arg1 & ")*(1+" & arg1 & ")"
      Else
        res = "=TANH(MMULT(TRANSPOSE(" & arg1 & ")," & arg2 & "))"
      End If
    Case "relu":
      If returnDerivativeFunction Then
        res = "IF(" & arg1 & ">0,1,0)"
      Else
        res = "=IF(0<MMULT(TRANSPOSE(" & arg1 & ")," & arg2 & ")," & "MMULT(TRANSPOSE(" & arg1 & ")," & arg2 & "),0)"
      End If
    Case "mlogit", "o:mlogit", "o:softmax", "softmax":
      If returnDerivativeFunction Then
        res = CVErr(xlErrNA)
      Else
        res = "=EXP(" & arg1 & ")/MMULT(TRANSPOSE(" & arg2 & "),EXP(" & arg1 & "))"
      End If
    Case "lin", "linear", "o:lin", "o:linear":
      If returnDerivativeFunction Then
        res = CVErr(xlErrNA)
      Else
        res = "=MMULT(TRANSPOSE(" & arg2 & ")," & arg1 & ")"
      End If
    Case "id", "o:id":
      If returnDerivativeFunction Then
        res = CVErr(xlErrNA)
      Else
        res = "=" & arg1
      End If
  End Select
  xlActfuncFormula = res
End Function


Function xlWeiGradFormula(ByVal strLossFunc As String, ByVal totNumLayers As Byte, ByVal layerLevel As Byte, Optional ws As Worksheet = Null)
  Dim res$, i&, j&, k&, deriv$, actfunc$
  If layerLevel <= totNumLayers + 1 And layerLevel >= 1 Then
    
    res = "(yhat-yobs)"
    actfunc = "logit"
    If Not IsNull(ws) Then actfunc = ws.Range("N_" & totNumLayers).Cells(2, 1)
    Select Case LCase(strLossFunc)
      Case "ls", "leastsquare":
        'res = "D_" & totNumLayers & "*(1-D_" & totNumLayers & ")*MMULT(fintrans," & res & ")"
        deriv = xlActfuncFormula(actfunc, "D_" & totNumLayers, , , True)
        res = deriv & "*MMULT(fintrans," & res & ")"
      
      Case "xen", "crossentropy":
        
        If layerLevel > totNumLayers Then GoTo e
        deriv = xlActfuncFormula(actfunc, "D_" & totNumLayers, , , True)
        res = deriv & "*" & res & ""
    
    End Select
    
    For k = totNumLayers - 1 To layerLevel Step -1
      If Not IsNull(ws) Then actfunc = ws.Range("N_" & k).Cells(2, 1)
      deriv = xlActfuncFormula(actfunc, "D_" & k, , , True)
      res = deriv & "*MMULT(W_" & k + 1 & "," & res & ")"
    Next k
    
    res = "=MMULT(D_" & layerLevel - 1 & ",TRANSPOSE(" & res & "))"
    
    xlWeiGradFormula = res: Exit Function
  End If
  
e:
  xlWeiGradFormula = "=NA()"
End Function

Sub train(ws As Worksheet)
On Error GoTo 0
  If DEBUG_LEVEL > 0 Then Application.ScreenUpdating = True: Application.EnableEvents = True
  Dim rngTrack As Range
  Dim timer As Double: timer = Now()
  Dim strMethod$: strMethod = ws.Range("method").Value
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
  Dim r As Range, c As Range, n As Long, i As Long, j As Long, loss_last As Double, loss_start As Double, loss_lastEpoch_train As Double, loss_lastEpoch_test As Double
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
lbl_run_new:
  Call initNextWeights(ws)

  Dim nrowsBatch As Long, iRoll As Long, nRoll As Long, iRepBatch As Long
  nRoll = Int((szTrain - szBatch) / roll)
  ws.Calculate
  While epoch > 0
    loss_lastEpoch_train = ws.Range("totloss").Value
    loss_lastEpoch_test = ws.Range("totloss_t").Value
lbl_rrmsprop_pre:
    Select Case strMethod
      Case "rrmsprop-":
        ws.Range("GradsPrevEpoch").Value2 = ws.Range("Grads").Value2
        ws.Range("method").Value = "rmsprop"
        ws.Calculate
    End Select
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
      ws.Names("D_" & j).RefersTo = "='" & ws.Name & "'!" & ws.Range("D_" & j & "i").Cells(1, 1).Resize(nrowsBatch, szBatch).AddressLocal
    Next j
    nrowsBatch = ws.Range("yhati").Rows.Count
    ws.Names("yhat").RefersTo = "='" & ws.Name & "'!" & ws.Range("yhati").Cells(1, 1).Resize(nrowsBatch, szBatch).AddressLocal
    nrowsBatch = ws.Range("yobsi").Rows.Count
    ws.Names("yobs").RefersTo = "='" & ws.Name & "'!" & ws.Range("yobsi").Cells(1, 1).Resize(nrowsBatch, szBatch).AddressLocal
    ws.Names("loss").RefersTo = "='" & ws.Name & "'!" & ws.Range("lossi").Cells(1, 1).Resize(1, szBatch).AddressLocal
    
    Dim nDropout As Long, zDropout()
    ReDim zDropout(1 To ws.Range("nWeights").Value, 1 To 2)
    

    
    For iRoll = 0 To nRoll Step 1
      loss_last = ws.Range("totloss").Value
      If DEBUG_LEVEL > 1 Then ws.Range("D_0").Select: Stop
      If DEBUG_LEVEL > 1 Then ws.Range("yobs").Select: Stop
      For iRepBatch = 1 To repBatch
lbl_update:
        ws.Calculate
        Application.StatusBar = "Epoch:" & epoch & " Batch:" & Format(iRoll, "000") & _
                " Step:" & Format(iRepBatch, "000") & " |Previous epoch's exit losses: test=" & _
                Format(loss_lastEpoch_test, "0.0000000000000000") & _
                " train=" & Format(loss_lastEpoch_train, "0.0000000000000000")

        ws.Range("prevState").Value2 = ws.Range("WorkRange").Value2
        ws.Range("Weights").Value2 = ws.Range("nextWeights").Value2
        For j = LBound(fmlacells, 1) To UBound(fmlacells, 1)
          ws.Range("Weights").Cells(fmlacells(j, 1), fmlacells(j, 2)).FormulaLocal = fmlacells(j, 3)
        Next j
  
lbl_dropout:
        If DO_DROPOUT Then
          j = 0
          For i = 1 To nLayers + 1
            If DEBUG_LEVEL > 2 Then ws.Range("W_" & i).Select: Stop
            For Each c In ws.Range("W_" & i).Cells
              If Left(Trim(c.FormulaLocal), 1) <> "=" Then
                If Rnd > 0.5 Then
                  j = j + 1
                  Set zDropout(j, 1) = c
                  zDropout(j, 2) = c.FormulaLocal
                  c.FormulaLocal = "=0"
                End If
              End If
            Next c
          Next i
          nDropout = j
        End If
lbl_post_update:
        Select Case ws.Range("method").Value
          Case "bp": 'nothing to do
          
          Case "rprop-": 'no weight backtracking
            
            ws.Range("prevRPROP").Value2 = ws.Range("rprop").Value2
            
          Case "rprop": 'has weight backtracking
            ws.Calculate
            If ws.Range("totloss") >= loss_last Then
              
              For Each r In ws.Range("Weights").Cells
                If Trim(r.Formula) <> "" And IsNumeric(r.Formula) Then
                  If Sgn(r.Offset(0, d).Value) <> Sgn(r.Offset(u, d).Value) Then
                    If DEBUG_LEVEL > 3 Then r.Select: Stop
                    r.Value = r.Offset(u, 0).Value
                  End If
                End If
              Next r
            End If
            If DEBUG_LEVEL > 1 Then ws.Range("prevRPROP").Select: Stop
            ws.Range("prevRPROP").Value2 = ws.Range("rprop").Value2
            If DEBUG_LEVEL > 1 Then ws.Range("prevRPROP").Select: Stop
          Case "rmsprop": 'decaying history
            ws.Calculate
            If DEBUG_LEVEL > 1 Then ws.Range("prevRMSPROP").Select: Stop
            ws.Range("prevRMSPROP").Value2 = ws.Range("rmsprop").Value2
            If DEBUG_LEVEL > 1 Then ws.Range("prevRMSPROP").Select: Stop
          Case Else:
          
        End Select
lbl_restore_dropout:
        If DO_DROPOUT Then
          For j = 1 To nDropout
            If DEBUG_LEVEL > 3 Then zDropout(j, 1).Select: Stop
            zDropout(j, 1).FormulaLocal = zDropout(j, 2)
            If DEBUG_LEVEL > 3 Then zDropout(j, 1).Select: Stop
          Next j
          ws.Calculate
        End If
        If EXIT_NOW Then GoTo lbl_SGD_roll_batch
      Next iRepBatch

lbl_SGD_roll_batch:
      nrowsBatch = ws.Range("D_0i").Rows.Count
      ws.Names("D_0").RefersTo = "='" & ws.Name & "'!" & ws.Range("D_0").Offset(0, roll).AddressLocal
      nrowsBatch = ws.Range("yobsi").Rows.Count
      ws.Names("yobs").RefersTo = "='" & ws.Name & "'!" & ws.Range("yobs").Offset(0, roll).AddressLocal
      
      If EXIT_NOW Then GoTo lbl_SGD_restore_fullbatch
    Next iRoll
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
lbl_rrmsprop_post:
    
    ws.Calculate
    Select Case strMethod
      Case "rrmsprop-":
        Dim baselr As Double: baselr = 0.1 / ws.Range("loss").Count * ws.Range("roll").Value
        Dim ang As Double: ang = Application.Evaluate("=ACOS(SUMPRODUCT(Grads,GradsPrevEpoch)/SQRT(SUMSQ(Grads))/SQRT(SUMSQ(GradsPrevEpoch)))/PI()*180")
        If Abs(ang) < 30 Then
          ws.Range("learningRate").Value = ws.Range("learningRate").Value * ws.Range("rpropup").Value
          If ws.Range("learningRate").Value > 5 * baselr Then ws.Range("learningRate").Value = baselr
        ElseIf Abs(ang) > 150 Then
          ws.Range("learningRate").Value = ws.Range("learningRate").Value * ws.Range("rpropdn").Value
          If ws.Range("learningRate").Value < 0.1 * baselr Then ws.Range("learningRate").Value = baselr
        Else
          ws.Range("learningRate").Value = baselr
        End If
        
    End Select
    epoch = epoch - 1
    
    Set rngTrack = Nothing
    Select Case ws.Range("DO_TRACKING").Value
      Case False:
      Case True, "loss":
        Set rngTrack = ws.Range("loss", "tloss")
      Case "CUSTOM":
        Set r = ws.Range("tracked").Offset(0, 1)
        If Trim(r.Formula) <> "" Then
          Set rngTrack = ws.Range(r, ws.Range("tracked").End(xlToRight))
        End If
      Case Else:
        If existRangeName(ws.Range("DO_TRACKING").Value, ws) Then
          Call prepareLinearizedTrackingRange(ws.Range("DO_TRACKING").Value, ws)
          Set r = ws.Range("tracked").Offset(0, 1)
          Set rngTrack = ws.Range(r, ws.Range("tracked").End(xlToRight))
        End If
    End Select
    If Not rngTrack Is Nothing Then
      Set r = shtTrace.Cells(100000, 3).End(xlUp).Offset(1, 0)
      r.Offset(0, -2).Value = Now()
      r.Offset(0, -1).Value = ws.Range("DO_TRACKING").Value
      r.Resize(1, rngTrack.Columns.Count).Value2 = rngTrack.Value2
    End If
    Select Case strMethod
      Case "rrmsprop-":
        ws.Range("method").Value = "rrmsprop-"
    End Select
    If DEBUG_LEVEL > 0 Then ws.Range("totloss").Select: Stop
    If EXIT_NOW Then GoTo lbl_msgbox
    
  Wend
lbl_msgbox:
  ws.Calculate
  Dim msg As String
  Select Case ws.Range("method").Value
    Case "bp":
      msg = "Backprop with learning rate " & lr
    Case "rprop-":
      msg = "rprop- (without weight backtracking) with learning rate resilience scalars {" & _
            ws.Range("rpropdn").Value & ", " & ws.Range("rpropup").Value & _
            "} and learning rate bounds [" & ws.Range("rpropfloor").Value & " to " & ws.Range("rpropcap").Value & "]"
    
    Case "rprop":
      msg = "rprop with learning rate resilience scalars {" & ws.Range("rpropdn").Value & ", " & _
            ws.Range("rpropup").Value & _
            "} and learning rate bounds [" & ws.Range("rpropfloor").Value & " to " & ws.Range("rpropcap").Value & "]"

    Case "rmsprop":
      msg = "rmsprop with global learning rate=" & ws.Range("learningRate").Value & ", mini batch size=" & _
            ws.Range("batch_size").Value & ", and roll=" & ws.Range("roll").Value
  End Select
  
  MsgBox "Trained " & ws.Range("epoch").Value & " epochs of " & msg & vbCr & "Time spent:  " & _
        Format(Now() - timer, "hh:mm:ss") & " (hh:mm:ss)" & vbCr & _
        "loss before = " & loss_start & vbCr & "loss now = " & ws.Range("totloss").Value
End Sub

