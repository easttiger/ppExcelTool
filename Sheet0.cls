VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  Dim xlcalcb4: xlcalcb4 = Application.Calculation
  If Target.Cells(1, 1).AddressLocal = Me.Range("train").AddressLocal Then
    Cancel = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo lbl_exit_run100i
    Call train(Me)
lbl_exit_run100i:
    Application.Calculation = xlcalcb4: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
  End If
  
  If Target.Cells(1, 1).AddressLocal = Me.Range("imputeNoise").AddressLocal Then
    Cancel = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo lbl_exit_imputeNoise
    Call imputeNoise(Me)
lbl_exit_imputeNoise:
    Application.Calculation = xlcalcb4: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
  End If
  
  If Target.Cells(1, 1).AddressLocal = Me.Range("trimWeiMats").AddressLocal Then
    Cancel = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo lbl_exit_trimWeiMats
    Call trimWeiMats(Me)
lbl_exit_trimWeiMats:
    Application.Calculation = xlcalcb4: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
  End If
  
  If Target.Cells(1, 1).AddressLocal = Me.Range("shrinkWeiMats").AddressLocal Then
    Cancel = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo lbl_exit_shrinkWeiMats
    Call shrinkWeiMats(Me)
lbl_exit_shrinkWeiMats:
    Application.Calculation = xlcalcb4: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
  End If
  
  If Target.Cells(1, 1).AddressLocal = Me.Range("plotNet").AddressLocal Then
    Cancel = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo lbl_exit_plotNet
    Call plotNetOnInstanceSheet(Me)
lbl_exit_plotNet:
    Application.Calculation = xlcalcb4: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
  End If
  
  If Target.Cells(1, 1).AddressLocal = Me.Range("removePlot").AddressLocal Then
    Cancel = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo lbl_exit_removePlot
    Call clearShapes(Me)
lbl_exit_removePlot:
    Application.Calculation = xlcalcb4: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
  End If
  
  If isContainedBy(Target.Cells(1, 1), Me.Range("Weights")) And Target.Cells(1, 1).Formula <> "" Then
    Cancel = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
On Error GoTo lbl_exit_selectConnector
    Me.Shapes("Conn_" & Replace(Target.AddressLocal, "$", "")).Select
lbl_exit_selectConnector:
    Application.Calculation = xlcalcb4: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
  End If
  
End Sub

