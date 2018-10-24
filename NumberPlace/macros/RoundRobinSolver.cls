VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RoundRobinSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Solver

'---------------------------------------
' メンバー変数宣言
'---------------------------------------
Private pInnerObj         As New Solver         '内部オブジェクト
Private pKnownNumber(9)   As String             '既知の数字
Private pMissingNumber(9) As String             '不足する数字
Private pPermPattern      As New Permutation    '順列管理用オブジェクト

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
    Dim rowIdx As Integer, colIdx As Integer
    Dim cell    As Range
    Dim findNum As String
    
    '盤面のRangeを内部オブジェクトに連携
    Call pInnerObj.Scan(iRange)
    
    '既知数配列、不足数配列の初期化（これらは行単位で管理）
    For rowIdx = 1 To 9
        pKnownNumber(rowIdx) = ""
        pMissingNumber(rowIdx) = "123456789"
    Next rowIdx
    
    '既知数と不足数のスキャン
    For rowIdx = 1 To 9
        For colIdx = 1 To 9
            Set cell = pInnerObj.Board(rowIdx, colIdx)
            findNum = cell.Value
            If findNum <> "" Then
                '盤面に数字が存在する場合
                pKnownNumber(rowIdx) = pKnownNumber(rowIdx) & findNum
                pMissingNumber(rowIdx) = Replace(pMissingNumber(rowIdx), findNum, "")
                
                '既知数の背景色をグレーにする
                If cell.Interior.Color = RGBEnum.White Then
                    cell.Interior.Color = RGBEnum.Silver
                End If
            Else
                '盤面に数字が存在しない場合
                cell.Interior.Color = RGBEnum.White
            End If
        Next colIdx
        
        'この行の不足数での順列を生成
        Call pPermPattern.GenPattern(rowIdx, pMissingNumber(rowIdx))
    Next rowIdx
End Sub

'=======================================
' 解答処理の実行
' @return True:配置完了
'=======================================
Private Function Solver_Execute() As Boolean
    ' まず、不要パターンを削除
    Call DelPattern

    ' 1行ずつパターン配置を試す
    Solver_Execute = TryPattern(1)
End Function

'---------------------------------------
' 列方向チェックで不要パターンの排除
'---------------------------------------
Private Sub DelPattern()
    Dim colIdx As Integer, rowIdx As Integer
    Dim boxRowIdx As Integer
    Dim pNo As Long
    Dim retCorrect As Boolean
    
    For rowIdx = 2 To 9
        For pNo = 1 To pPermPattern.Count(rowIdx)
            '順列パターンの仮配置
            If PlaceLine(rowIdx, pNo) Then
                '列方向の重複チェック
                For colIdx = 1 To 9
                    retCorrect = pInnerObj.IsCorrectVert(colIdx, rowIdx, pNo)
                    If Not retCorrect Then
                        Call pPermPattern.DelPattern(rowIdx, pNo)
                        Exit For
                    End If
                Next colIdx
                
                If retCorrect Then
                    'ボックスでの重複チェック
                    boxRowIdx = Fix((rowIdx - 1) / 3) * 3 + 1
                    For colIdx = 1 To 9 Step 3
                        If Not pInnerObj.IsCorrectBox(boxRowIdx, colIdx, rowIdx, pNo) Then
                            Call pPermPattern.DelPattern(rowIdx, pNo)
                            Exit For
                        End If
                    Next colIdx
                End If
            End If
        Next pNo
        
        '順列パターンの不要箇所をシュリンクアップする
        Call pPermPattern.Shrink(rowIdx)
        '仮配置したパターンをリセットする
        Call ResetLine(rowIdx)
    Next rowIdx
End Sub

'---------------------------------------
' 1行ずつ順列総当りでパターン配置と確認を繰り返す
'
' @return True:配置完了
'---------------------------------------
Private Function TryPattern(iRowIdx As Integer) As Boolean
    Dim pNo As Long
    
    TryPattern = False
    DoEvents
    
    '不足数の仮配置と検証を、順列のパターン数回繰り返す
    For pNo = 1 To pPermPattern.Count(iRowIdx)
        '順列パターンの仮配置
        If PlaceLine(iRowIdx, pNo) Then
            '検証
            If pInnerObj.IsCorrect(iRowIdx, pNo) Then
                '9行目で検証ＯＫなら完成。
                If iRowIdx >= 9 Then
                    TryPattern = True
                    Exit Function
                End If
                
                '9行目でない場合は、次の行で試行
                TryPattern = TryPattern(iRowIdx + 1)
                If TryPattern Then
                    'True復帰なら完成
                    Exit Function
                End If
            End If
        End If
    Next pNo
    
    'ループを回り切った場合、検証ＮＧ。
    '仮配置したパターンをリセットする
    Call ResetLine(iRowIdx)
End Function

'---------------------------------------
' 数字の配置（行方向）
'
' @return True:配置を行った、False:配置しなかった
'---------------------------------------
Private Function PlaceLine(iRowIdx As Integer, iPatternNo As Long) As Boolean
    Dim numPattern As String
    Dim cell As Range
    Dim colIdx As Integer
    
    numPattern = pPermPattern.GetPattern(iRowIdx, iPatternNo)
    If numPattern = "" Then
        PlaceLine = False
        Exit Function
    End If
    
    For colIdx = 1 To 9
        Set cell = pInnerObj.Board(iRowIdx, colIdx)
        If cell.Interior.Color = RGBEnum.White Then
            cell.Value = left(numPattern, 1)
            numPattern = mid(numPattern, 2)
        End If
    Next colIdx
    
    PlaceLine = True
End Function

'---------------------------------------
' 数字の配置のリセット
'---------------------------------------
Private Sub ResetLine(iRowIdx As Integer)
    Dim cell As Range
    Dim colIdx As Integer
    
    For colIdx = 1 To 9
        Set cell = pInnerObj.Board(iRowIdx, colIdx)
        If cell.Interior.Color = RGBEnum.White Then
            cell.Value = ""
        End If
    Next colIdx
End Sub

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

