VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Solver

'---------------------------------------
' メンバー変数宣言
'---------------------------------------
Private pInnerObj As New Solver             '内部オブジェクト

'=======================================
' 盤面範囲オブジェクトの返却
'=======================================
Private Property Get Solver_Board() As Range
    Set Solver_Board = pInnerObj.Board
End Property

'=======================================
' 処理中の盤面をリセットする
'=======================================
Private Sub Solver_Clear(iRange As Range)
    Call pInnerObj.Clear(iRange)
End Sub

'=======================================
' 盤面のスキャン
'=======================================
Private Sub Solver_Scan(iRange As Variant)
    '盤面のRangeを内部オブジェクトに連携
    Call pInnerObj.Scan(iRange)
End Sub

'=======================================
' 解答処理の実行
' @return True:配置完了
'=======================================
Private Function Solver_Execute() As Boolean
    Solver_Execute = pInnerObj.Execute()
End Function

'=======================================
' 盤面全体を埋めることができたか？
'=======================================
Private Function Solver_IsFilled() As Boolean
    Solver_IsFilled = pInnerObj.IsFilled
End Function

'=======================================
' 盤面全体で数字の配置が完成しているか？
'=======================================
Private Function Solver_IsCorrect(iRowIdx As Integer, iPNo As Long) As Boolean
    Solver_IsCorrect = pInnerObj.IsCorrect(iRowIdx, iPNo)
End Function

'=======================================
' 単一ブロック（枠）で配置が完成しているか？
'=======================================
Private Function Solver_IsCorrectBox(iRow As Integer, iCol As Integer, iRowIdx As Integer, iPNo As Long) As Boolean
    Solver_IsCorrectBox = pInnerObj.IsCorrectBox(iRow, iCol, iRowIdx, iPNo)
End Function

'=======================================
' 単一ブロック（縦）で配置が完成しているか？
'=======================================
Private Function Solver_IsCorrectVert(iCol As Integer, iRowIdx As Integer, iPNo As Long) As Boolean
    Solver_IsCorrectVert = pInnerObj.IsCorrectVert(iCol, iRowIdx, iPNo)
End Function

