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
' �����o�[�ϐ��錾
'---------------------------------------
Private pInnerObj         As New Solver         '�����I�u�W�F�N�g
Private pKnownNumber(9)   As String             '���m�̐���
Private pMissingNumber(9) As String             '�s�����鐔��
Private pPermPattern      As New Permutation    '����Ǘ��p�I�u�W�F�N�g

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
    
    '�Ֆʂ�Range������I�u�W�F�N�g�ɘA�g
    Call pInnerObj.Scan(iRange)
    
    '���m���z��A�s�����z��̏������i�����͍s�P�ʂŊǗ��j
    For rowIdx = 1 To 9
        pKnownNumber(rowIdx) = ""
        pMissingNumber(rowIdx) = "123456789"
    Next rowIdx
    
    '���m���ƕs�����̃X�L����
    For rowIdx = 1 To 9
        For colIdx = 1 To 9
            Set cell = pInnerObj.Board(rowIdx, colIdx)
            findNum = cell.Value
            If findNum <> "" Then
                '�Ֆʂɐ��������݂���ꍇ
                pKnownNumber(rowIdx) = pKnownNumber(rowIdx) & findNum
                pMissingNumber(rowIdx) = Replace(pMissingNumber(rowIdx), findNum, "")
                
                '���m���̔w�i�F���O���[�ɂ���
                If cell.Interior.Color = RGBEnum.White Then
                    cell.Interior.Color = RGBEnum.Silver
                End If
            Else
                '�Ֆʂɐ��������݂��Ȃ��ꍇ
                cell.Interior.Color = RGBEnum.White
            End If
        Next colIdx
        
        '���̍s�̕s�����ł̏���𐶐�
        Call pPermPattern.GenPattern(rowIdx, pMissingNumber(rowIdx))
    Next rowIdx
End Sub

'=======================================
' �𓚏����̎��s
' @return True:�z�u����
'=======================================
Private Function Solver_Execute() As Boolean
    ' �܂��A�s�v�p�^�[�����폜
    Call DelPattern

    ' 1�s���p�^�[���z�u������
    Solver_Execute = TryPattern(1)
End Function

'---------------------------------------
' ������`�F�b�N�ŕs�v�p�^�[���̔r��
'---------------------------------------
Private Sub DelPattern()
    Dim colIdx As Integer, rowIdx As Integer
    Dim boxRowIdx As Integer
    Dim pNo As Long
    Dim retCorrect As Boolean
    
    For rowIdx = 2 To 9
        For pNo = 1 To pPermPattern.Count(rowIdx)
            '����p�^�[���̉��z�u
            If PlaceLine(rowIdx, pNo) Then
                '������̏d���`�F�b�N
                For colIdx = 1 To 9
                    retCorrect = pInnerObj.IsCorrectVert(colIdx, rowIdx, pNo)
                    If Not retCorrect Then
                        Call pPermPattern.DelPattern(rowIdx, pNo)
                        Exit For
                    End If
                Next colIdx
                
                If retCorrect Then
                    '�{�b�N�X�ł̏d���`�F�b�N
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
        
        '����p�^�[���̕s�v�ӏ����V�������N�A�b�v����
        Call pPermPattern.Shrink(rowIdx)
        '���z�u�����p�^�[�������Z�b�g����
        Call ResetLine(rowIdx)
    Next rowIdx
End Sub

'---------------------------------------
' 1�s�����񑍓���Ńp�^�[���z�u�Ɗm�F���J��Ԃ�
'
' @return True:�z�u����
'---------------------------------------
Private Function TryPattern(iRowIdx As Integer) As Boolean
    Dim pNo As Long
    
    TryPattern = False
    DoEvents
    
    '�s�����̉��z�u�ƌ��؂��A����̃p�^�[������J��Ԃ�
    For pNo = 1 To pPermPattern.Count(iRowIdx)
        '����p�^�[���̉��z�u
        If PlaceLine(iRowIdx, pNo) Then
            '����
            If pInnerObj.IsCorrect(iRowIdx, pNo) Then
                '9�s�ڂŌ��؂n�j�Ȃ犮���B
                If iRowIdx >= 9 Then
                    TryPattern = True
                    Exit Function
                End If
                
                '9�s�ڂłȂ��ꍇ�́A���̍s�Ŏ��s
                TryPattern = TryPattern(iRowIdx + 1)
                If TryPattern Then
                    'True���A�Ȃ犮��
                    Exit Function
                End If
            End If
        End If
    Next pNo
    
    '���[�v�����؂����ꍇ�A���؂m�f�B
    '���z�u�����p�^�[�������Z�b�g����
    Call ResetLine(iRowIdx)
End Function

'---------------------------------------
' �����̔z�u�i�s�����j
'
' @return True:�z�u���s�����AFalse:�z�u���Ȃ�����
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
' �����̔z�u�̃��Z�b�g
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
