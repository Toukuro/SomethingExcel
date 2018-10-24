VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EliminationSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Solver

'---------------------------------------
' メンバー変数宣言
'---------------------------------------
Private pInnerObj As New Solver             '内部オブジェクト
Private pNumPositions(9) As Collection      '盤面上の各数字の位置

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
    Dim pos     As Position
    
    '盤面のRangeを内部オブジェクトに連携
    Call pInnerObj.Scan(iRange)
    
    '数字の位置を保持する配列にCollectionオブジェクトを設定
    For rowIdx = 1 To 9
        Set pNumPositions(rowIdx) = New Collection
    Next rowIdx
    
    For rowIdx = 1 To 9
        For colIdx = 1 To 9
            Set cell = pInnerObj.Board(rowIdx, colIdx)
            findNum = cell.Value
            If findNum <> "" Then
                Set pos = New Position
                Call pos.SetXYbyInt(colIdx, rowIdx)
                Call pNumPositions(Val(findNum)).Add(pos)
                
                If cell.Interior.Color = RGBEnum.White Then
                    cell.Interior.Color = RGBEnum.Silver
                End If
            End If
        Next colIdx
    Next rowIdx
End Sub

'=======================================
' 解答処理の実行
' @return True:配置完了
'=======================================
Private Function Solver_Execute() As Boolean
    Dim canTry As Boolean
    Dim num As Integer
    
    Do
        canTry = False
        For num = 1 To 9
            canTry = canTry Or PlaceNumber(num)
        Next num
    Loop While canTry
    
    Solver_Execute = pInnerObj.IsFilled
End Function

'---------------------------------------
' 数字の配置
'
' @return 数字を配置することができた場合にTrue
'---------------------------------------
Private Function PlaceNumber(iNum As Integer) As Boolean
    Dim cnt As Integer
    Dim rowIdx As Integer, colIdx As Integer
    
    PlaceNumber = False
    
    '記録した位置の数が９なら、この数字は完成なので処理をスキップ
    If pNumPositions(iNum).Count >= 9 Then
        Exit Function
    End If
    
    '指定位置より上下左右に同一番号の配置不可マスクを設定
    For cnt = 1 To pNumPositions(iNum).Count
        Call SetMask(pNumPositions(iNum)(cnt))
    Next cnt
    
    For rowIdx = 1 To 9 Step 3
        For colIdx = 1 To 9 Step 3
            PlaceNumber = PlaceNumber Or TryPlace(iNum, rowIdx, colIdx)
        Next colIdx
    Next rowIdx
    
    '配置不可マスクを解除
    Call ResetMask
End Function

'---------------------------------------
' 指定位置の上下左右に配置不可のマスクを設定
'---------------------------------------
Private Sub SetMask(iPos As Position)
    Dim cellInterior As Interior
    Dim idx As Integer
    
    For idx = 1 To 9
        '行方向
        Set cellInterior = pInnerObj.Board(iPos.IntY, idx).Interior
        If cellInterior.Color = RGBEnum.White Then
            cellInterior.Color = RGBEnum.Red
        End If
        
        '列方向
        Set cellInterior = pInnerObj.Board(idx, iPos.IntX).Interior
        If cellInterior.Color = RGBEnum.White Then
            cellInterior.Color = RGBEnum.Red
        End If
    Next idx
End Sub

'---------------------------------------
' 配置不可マスクを解除
'---------------------------------------
Private Sub ResetMask()
    Dim cellInterior As Interior
    Dim rowIdx As Integer
    Dim colIdx As Integer
    
    For rowIdx = 1 To 9
        For colIdx = 1 To 9
            Set cellInterior = pInnerObj.Board(rowIdx, colIdx).Interior
            If cellInterior.Color = RGBEnum.Red Then
                cellInterior.Color = RGBEnum.White
            End If
        Next colIdx
    Next rowIdx
End Sub

'---------------------------------------
' ボックス毎の数字の配置
'---------------------------------------
Private Function TryPlace(iNum As Integer, iRow As Integer, iCol As Integer) As Boolean
    Dim dRow As Integer, dCol As Integer
    Dim cell As Range
    Dim pos As New Position
    Dim cnt As Integer
    
    TryPlace = False
    
    'まず、このボックス内にiNumと同じ数字が無いことを確認
    For dRow = 0 To 2
        For dCol = 0 To 2
            Set cell = pInnerObj.Board(iRow + dRow, iCol + dCol)
            If cell.Value = CStr(iNum) Then
                Exit Function
            End If
        Next dCol
    Next dRow
    
    'ボックス内の白地セルのカウントと位置の保持
    cnt = 0
    For dRow = 0 To 2
        For dCol = 0 To 2
            Set cell = pInnerObj.Board(iRow + dRow, iCol + dCol)
            If cell.Interior.Color = RGBEnum.White Then
                Call pos.SetXYbyInt(iCol + dCol, iRow + dRow)
                cnt = cnt + 1
                '白地セル数が2以上ならアウト！
                If cnt > 1 Then Exit Function
            End If
        Next dCol
    Next dRow
    
    '白地セルが１個だけだった場合、数字の配置と位置をCollectionに追加
    If cnt = 1 Then
        Set cell = pInnerObj.Board(pos.IntY, pos.IntX)
        cell.Value = CStr(iNum)
        cell.Interior.Color = RGBEnum.Silver2
        Call pNumPositions(Val(iNum)).Add(pos)
        
        TryPlace = True
    End If
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

