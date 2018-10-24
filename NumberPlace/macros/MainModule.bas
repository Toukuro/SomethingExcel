Attribute VB_Name = "MainModule"
'---------------------------------------
' グローバル変数
'---------------------------------------
' RGB Color Definitions (BBGGRR)
Public Enum RGBEnum
    White = &HFFFFFF
    Black = &H0
    Red = &HFF
    Silver = &HC0C0C0
    Silver2 = &HE0E0E0
End Enum

'---------------------------------------
' 解答実行
'---------------------------------------
Public Sub Solving()
    Dim SolverCollection As New Collection
    Dim solv As Solver
    Dim numRange As Range
    Dim i As Integer
    
    DisableAutomatic
    
    '盤面の範囲
    Set numRange = ActiveSheet.Range(Cells(2, 2), Cells(10, 10))
    
    '解法パターンをコレクションに設定
    SolverCollection.Add New EliminationSolver
    SolverCollection.Add New RoundRobinSolver
    
    'すべての解法を使って問題を解く
    For i = 1 To SolverCollection.Count
        Set solv = SolverCollection(i)
        Call solv.Scan(numRange)
        If solv.Execute Then
            MsgBox "解答完了"
            EnableAutomatic
            Exit Sub
        End If
    Next i
    
    MsgBox "解答失敗"
    EnableAutomatic
End Sub

'---------------------------------------
' 盤面リセット
'---------------------------------------
Public Sub Clear()
    Dim solv As New Solver
    Dim numRange As Range
    
    Set numRange = ActiveSheet.Range(Cells(2, 2), Cells(10, 10))
    Call solv.Clear(numRange)
    
    EnableAutomatic
End Sub

'---------------------------------------
' 自動計算の無効化
'---------------------------------------
Private Sub DisableAutomatic()
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False
End Sub

'---------------------------------------
' 自動計算の有効化
'---------------------------------------
Private Sub EnableAutomatic()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
