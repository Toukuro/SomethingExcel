Attribute VB_Name = "Module1"
'-------- teaching 0
Sub OK0_Click()
    Call SelectedAdjust("AG4:AG19", "AH4:AH19", 0.01, 4, -2)
End Sub

Sub NG0_Click()
    Call SelectedAdjust("AG4:AG19", "AH4:AH19", -0.01, 4, -2)
End Sub

'-------- teaching 1
Sub OK1_Click()
    Call SelectedAdjust("AL4:AL19", "AM4:AM19", 0.01, 4, -2)
End Sub

Sub NG1_Click()
    Call SelectedAdjust("AL4:AL19", "AM4:AM19", -0.01, 4, -2)
End Sub

'-------- teaching 2
Sub OK2_Click()
    Call SelectedAdjust("AQ4:AQ19", "AR4:AR19", 0.01, 4, -2)
End Sub

Sub NG2_Click()
    Call SelectedAdjust("AQ4:AQ19", "AR4:AR19", -0.01, 4, -2)
End Sub

'-------- teaching 3
Sub OK3_Click()
    Call SelectedAdjust("AV4:AV19", "AW4:AW19", 0.01, 4, -2)
End Sub

Sub NG3_Click()
    Call SelectedAdjust("AV4:AV19", "AW4:AW19", -0.01, 4, -2)
End Sub

'-------- teaching 4
Sub OK4_Click()
    Call SelectedAdjust("BA4:BA19", "BB4:BB19", 0.01, 4, -2)
End Sub

Sub NG4_Click()
    Call SelectedAdjust("BA4:BA19", "BB4:BB19", -0.01, 4, -2)
End Sub


'-------- teaching 5
Sub OK5_Click()
    Call SelectedAdjust("AG24:AG39", "AH24:AH39", 0.01, 4, -2)
End Sub

Sub NG5_Click()
    Call SelectedAdjust("AG24:AG39", "AH24:AH39", -0.01, 4, -2)
End Sub

'-------- teaching 6
Sub OK6_Click()
    Call SelectedAdjust("AL24:AL39", "AM24:AM39", 0.01, 4, -2)
End Sub

Sub NG6_Click()
    Call SelectedAdjust("AL24:AL39", "AM24:AM39", -0.01, 4, -2)
End Sub

'-------- teaching 7
Sub OK7_Click()
    Call SelectedAdjust("AQ24:AQ39", "AR24:AR39", 0.01, 4, -2)
End Sub

Sub NG7_Click()
    Call SelectedAdjust("AQ24:AQ39", "AR24:AR39", -0.01, 4, -2)
End Sub

'-------- teaching 8
Sub OK8_Click()
    Call SelectedAdjust("AV24:AV39", "AW24:AW39", 0.01, 4, -2)
End Sub

Sub NG8_Click()
    Call SelectedAdjust("AV24:AV39", "AW24:AW39", -0.01, 4, -2)
End Sub

'-------- teaching 9
Sub OK9_Click()
    Call SelectedAdjust("BA24:BA39", "BB24:BB39", 0.01, 4, -2)
End Sub

Sub NG9_Click()
    Call SelectedAdjust("BA24:BA39", "BB24:BB39", -0.01, 4, -2)
End Sub


Private Sub SelectedAdjust(iRange1 As String, iRange2 As String, iDelta As Double, Optional iMax As Double = 1#, Optional iMin As Double = 0#)
    Dim r1 As Range, r2 As Range
    Dim i As Integer
    
    Set r1 = ActiveSheet.Range(iRange1)
    Set r2 = ActiveSheet.Range(iRange2)
    
    For i = 1 To r1.Rows.Count
        If r1(i, 1) >= 1# Then
            Dim newVal As Double
            
            newVal = r2(i, 1) + iDelta
            If newVal > iMax Then newVal = iMax
            If newVal < iMin Then newVal = iMin
            
            r2(i, 1) = newVal
        End If
    Next
End Sub


