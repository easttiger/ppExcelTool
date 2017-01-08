VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cbtn_charting_Click()
On Error GoTo e
  Const LEFTPOS& = 10
  Const CHART_WIDTH& = 500
  
    Dim ch1 As Chart, ch2 As Chart, r1 As Range, r2 As Range, ans
    Dim maxrow&, maxcol&
    maxrow = Me.Cells(20000, 2).End(xlUp).Row - 1
    maxcol = Me.Cells(2, 10000).End(xlToLeft).Column - 2
    Dim finish As Boolean
    Dim ncharts As Long: ncharts = 0
    While Not finish
      ans = InputBox("Enter" & vbNewLine & "   Start-cell-address,nrow,ncol" & vbNewLine & "comma separated or Enter ""finish"" to finish:", , "C2,100," & shtPrep.Range("ntrain"))
      If Trim(LCase(ans)) = "finish" Then GoTo e
      ans = Split(ans, ",")
      Dim r0 As Range: Set r0 = Me.Range(ans(LBound(ans)))
      Dim nrow As Long: nrow = CLng(ans(LBound(ans) + 1)): nrow = Application.Min(maxrow, nrow)
      Dim ncol As Long: ncol = CLng(ans(LBound(ans) + 2)): ncol = Application.Min(maxcol, ncol)
      
      Set r1 = r0.Resize(nrow, ncol)
      r1.Select
      shtTrace.Shapes.AddChart2(227, xlLine).Select
      
      Set ch1 = ActiveChart
      With ch1
        .SetSourceData Source:=Range("log!" & r1.AddressLocal)
        .Legend.Delete
        .ChartTitle.Text = InputBox("Chart Title:", , "Training Errors")
        .ChartArea.Select
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 40
        .Axes(xlCategory).TickLabels.Font.Size = 20
        .Axes(xlValue).TickLabels.Font.Size = 20
        With .Parent
          .Height = CHART_WIDTH
          .Width = CHART_WIDTH
          .Top = 100
          .Left = LEFTPOS& + (CHART_WIDTH + 3) * ncharts
        End With
        .PlotBy = xlColumns
      End With
      ncharts = ncharts + 1
    Wend
    
e:
End Sub


Function makeChart(ByVal dataspec As String, ByVal titlespec As String, ByVal lngTop As Long, ByVal lngLeft As Long)
On Error GoTo e
  Const LEFTPOS& = 10
  Const CHART_WIDTH& = 500
  
    Dim ch1 As Chart, ch2 As Chart, r1 As Range, r2 As Range, ans
    Dim maxrow&, maxcol&
    maxrow = Me.Cells(200000, 2).End(xlUp).Row - 1
    maxcol = Me.Cells(2, 10000).End(xlToLeft).Column - 2

      If Trim(LCase(dataspec)) = "finish" Then GoTo e
      ans = Split(dataspec, ",")
      Dim r0 As Range: Set r0 = Me.Range(ans(LBound(ans)))
      Dim nrow As Long: nrow = CLng(ans(LBound(ans) + 1)): nrow = Application.Min(maxrow, nrow)
      Dim ncol As Long: ncol = CLng(ans(LBound(ans) + 2)): ncol = Application.Min(maxcol, ncol)
      
      Set r1 = r0.Resize(nrow, ncol)
      r1.Select
      shtTrace.Shapes.AddChart2(227, xlLine).Select
      
      Set ch1 = ActiveChart
      With ch1
        .SetSourceData Source:=Range("log!" & r1.AddressLocal)
        .Legend.Delete
        .ChartTitle.Text = titlespec
        .ChartArea.Select
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 40
        .Axes(xlCategory).TickLabels.Font.Size = 20
        .Axes(xlValue).TickLabels.Font.Size = 20
        With .Parent
          .Height = CHART_WIDTH
          .Width = CHART_WIDTH
          .Top = lngTop
          .Left = lngLeft
        End With
        .PlotBy = xlColumns
      End With
    ch1.ExportAsFixedFormat xlTypePDF, Replace(shtDoe.Range("save_folder").Value & ch1.ChartTitle.Text & ".pdf", " ", "")
e:
End Function
