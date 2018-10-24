VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Solver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------
' メンバー変数
'---------------------------------------
Private pBoard As Range

'=======================================
' 盤面範囲オブジェクトの返却
'=======================================
Public Property Get Board() As Range
    Set Board = pBoard
End Property

'=======================================
' 処理中の盤面をリセットする
'=======================================
Public Sub Clear(iRange As Range)
    Dim rowIdx As Integer, colIdx As Integer
    
    Set pBoard = iRange
    For rowIdx = 1 To 9
        For colIdx = 1 To 9
            Dim cell As Range
            Set cell = pBoard(rowIdx, colIdx)
            
            If cell.Interior.Color = RGBEnum.Silver Then
                cell.Interior.Color = RGBEnum.White
            Else
                cell.Value = ""
                cell.Interior.Color = RGBEnum.White
            End If
        Next colIdx
        DoEvents
    Next rowIdx
End Sub

'=======================================
' 盤面のスキャン
'=======================================
Public Sub Scan(iRange)
    Set pBoard = iRange
End Sub

'=======================================
' 解答処理の実行
' @return True:配置完了
'=======================================
Public Function Execute() As Boolean
    Execute = False
End Function

'=======================================
' 盤面全体を埋めることができたか？
'=======================================
Public Function IsFilled() As Boolean
    Dim rowIdx As Integer, colIdx As Integer
    
    IsFilled = True
    
    For rowIdx = 1 To 9
        For colIdx = 1 To 9
            If pBoard(rowIdx, colIdx).Value = "" Then
                IsFilled = False
                Exit Function
            End If
        Next colIdx
    Next rowIdx
    
End Function

'=======================================
' 盤面全体で数字の配置が完成しているか？
'=======================================
Public Function IsCorrect(iRowIdx As Integer, iPNo As Long) As Boolean
    Dim rowIdx As Integer, colIdx As Integer
    
    IsCorrect = False
    
    With pBoard
        '縦方向チェック
        For colIdx = 1 To 9
            If Not IsCorrectVert(colIdx, iRowIdx, iPNo) Then
                Exit Function
            End If
        Next colIdx
        
        'ブロックチェック
        For rowIdx = 1 To 9 Step 3
            For colIdx = 1 To 9 Step 3
                If Not IsCorrectBox(rowIdx, colIdx, iRowIdx, iPNo) Then
                    Exit Function
                End If
            Next colIdx
        Next rowIdx
    End With
    
    IsCorrect = True
End Function

'=======================================
' 単一ブロック（枠）で配置が完成しているか？
'=======================================
Public Function IsCorrectBox(iRow As Integer, iCol As Integer, iRowIdx As Integer, iPNo As Long) As Boolean
    Dim placedNum As String
    Dim dRow As Integer, dCol As Integer
    Dim usedNum As String
    
    IsCorrectBox = False
        
    usedNum = ""
    For dRow = 0 To 2
        For dCol = 0 To 2
            placedNum = pBoard(iRow + dRow, iCol + dCol).Value
            
            If placedNum <> "" Then
                If InStr(usedNum, placedNum) > 0 Then
                    Exit Function
                End If
                usedNum = usedNum & placedNum
            End If
        Next dCol
    Next dRow
    
    IsCorrectBox = True
End Function

'=======================================
' 単一ブロック（縦）で配置が完成しているか？
'=======================================
Public Function IsCorrectVert(iCol As Integer, iRowIdx As Integer, iPNo As Long) As Boolean
    Dim placedNum As String
    Dim rowIdx As Integer
    Dim usedNum As String
    
    IsCorrectVert = False
    
    '指定列で数字の重複があればNG
    usedNum = ""
    For rowIdx = 1 To 9
        placedNum = pBoard(rowIdx, iCol).Value
        
        If placedNum <> "" Then
            If InStr(usedNum, placedNum) > 0 Then
                Exit Function
            End If
            usedNum = usedNum & placedNum
        End If
    Next rowIdx
    
    IsCorrectVert = True
End Function

