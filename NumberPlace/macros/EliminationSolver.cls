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
' �����o�[�ϐ��錾
'---------------------------------------
Private pInnerObj As New Solver             '�����I�u�W�F�N�g
Private pNumPositions(9) As Collection      '�Ֆʏ�̊e�����̈ʒu

'=======================================
' �Ֆʔ͈̓I�u�W�F�N�g�̕ԋp
'=======================================
Private Property Get Solver_Board() As Range
    Set Solver_Board = pInnerObj.Board
End Property

'=======================================
' �������̔Ֆʂ����Z�b�g����
'=======================================
Private Sub Solver_Clear(iRange As Range)
    Call pInnerObj.Clear(iRange)
End Sub

'=======================================
' �Ֆʂ̃X�L����
'=======================================
Private Sub Solver_Scan(iRange As Variant)
    Dim rowIdx As Integer, colIdx As Integer
    Dim cell    As Range
    Dim findNum As String
    Dim pos     As Position
    
    '�Ֆʂ�Range������I�u�W�F�N�g�ɘA�g
    Call pInnerObj.Scan(iRange)
    
    '�����̈ʒu��ێ�����z���Collection�I�u�W�F�N�g��ݒ�
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
' �𓚏����̎��s
' @return True:�z�u����
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
' �����̔z�u
'
' @return ������z�u���邱�Ƃ��ł����ꍇ��True
'---------------------------------------
Private Function PlaceNumber(iNum As Integer) As Boolean
    Dim cnt As Integer
    Dim rowIdx As Integer, colIdx As Integer
    
    PlaceNumber = False
    
    '�L�^�����ʒu�̐����X�Ȃ�A���̐����͊����Ȃ̂ŏ������X�L�b�v
    If pNumPositions(iNum).Count >= 9 Then
        Exit Function
    End If
    
    '�w��ʒu���㉺���E�ɓ���ԍ��̔z�u�s�}�X�N��ݒ�
    For cnt = 1 To pNumPositions(iNum).Count
        Call SetMask(pNumPositions(iNum)(cnt))
    Next cnt
    
    For rowIdx = 1 To 9 Step 3
        For colIdx = 1 To 9 Step 3
            PlaceNumber = PlaceNumber Or TryPlace(iNum, rowIdx, colIdx)
        Next colIdx
    Next rowIdx
    
    '�z�u�s�}�X�N������
    Call ResetMask
End Function

'---------------------------------------
' �w��ʒu�̏㉺���E�ɔz�u�s�̃}�X�N��ݒ�
'---------------------------------------
Private Sub SetMask(iPos As Position)
    Dim cellInterior As Interior
    Dim idx As Integer
    
    For idx = 1 To 9
        '�s����
        Set cellInterior = pInnerObj.Board(iPos.IntY, idx).Interior
        If cellInterior.Color = RGBEnum.White Then
            cellInterior.Color = RGBEnum.Red
        End If
        
        '�����
        Set cellInterior = pInnerObj.Board(idx, iPos.IntX).Interior
        If cellInterior.Color = RGBEnum.White Then
            cellInterior.Color = RGBEnum.Red
        End If
    Next idx
End Sub

'---------------------------------------
' �z�u�s�}�X�N������
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
' �{�b�N�X���̐����̔z�u
'---------------------------------------
Private Function TryPlace(iNum As Integer, iRow As Integer, iCol As Integer) As Boolean
    Dim dRow As Integer, dCol As Integer
    Dim cell As Range
    Dim pos As New Position
    Dim cnt As Integer
    
    TryPlace = False
    
    '�܂��A���̃{�b�N�X����iNum�Ɠ����������������Ƃ��m�F
    For dRow = 0 To 2
        For dCol = 0 To 2
            Set cell = pInnerObj.Board(iRow + dRow, iCol + dCol)
            If cell.Value = CStr(iNum) Then
                Exit Function
            End If
        Next dCol
    Next dRow
    
    '�{�b�N�X���̔��n�Z���̃J�E���g�ƈʒu�̕ێ�
    cnt = 0
    For dRow = 0 To 2
        For dCol = 0 To 2
            Set cell = pInnerObj.Board(iRow + dRow, iCol + dCol)
            If cell.Interior.Color = RGBEnum.White Then
                Call pos.SetXYbyInt(iCol + dCol, iRow + dRow)
                cnt = cnt + 1
                '���n�Z������2�ȏ�Ȃ�A�E�g�I
                If cnt > 1 Then Exit Function
            End If
        Next dCol
    Next dRow
    
    '���n�Z�����P�����������ꍇ�A�����̔z�u�ƈʒu��Collection�ɒǉ�
    If cnt = 1 Then
        Set cell = pInnerObj.Board(pos.IntY, pos.IntX)
        cell.Value = CStr(iNum)
        cell.Interior.Color = RGBEnum.Silver2
        Call pNumPositions(Val(iNum)).Add(pos)
        
        TryPlace = True
    End If
End Function

'=======================================
' �ՖʑS�̂𖄂߂邱�Ƃ��ł������H
'=======================================
Private Function Solver_IsFilled() As Boolean
    Solver_IsFilled = pInnerObj.IsFilled
End Function

'=======================================
' �ՖʑS�̂Ő����̔z�u���������Ă��邩�H
'=======================================
Private Function Solver_IsCorrect(iRowIdx As Integer, iPNo As Long) As Boolean
    Solver_IsCorrect = pInnerObj.IsCorrect(iRowIdx, iPNo)
End Function

'=======================================
' �P��u���b�N�i�g�j�Ŕz�u���������Ă��邩�H
'=======================================
Private Function Solver_IsCorrectBox(iRow As Integer, iCol As Integer, iRowIdx As Integer, iPNo As Long) As Boolean
    Solver_IsCorrectBox = pInnerObj.IsCorrectBox(iRow, iCol, iRowIdx, iPNo)
End Function

'=======================================
' �P��u���b�N�i�c�j�Ŕz�u���������Ă��邩�H
'=======================================
Private Function Solver_IsCorrectVert(iCol As Integer, iRowIdx As Integer, iPNo As Long) As Boolean
    Solver_IsCorrectVert = pInnerObj.IsCorrectVert(iCol, iRowIdx, iPNo)
End Function
