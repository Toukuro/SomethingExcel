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
' �����o�[�ϐ��錾
'---------------------------------------
Private pInnerObj As New Solver             '�����I�u�W�F�N�g

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
    '�Ֆʂ�Range������I�u�W�F�N�g�ɘA�g
    Call pInnerObj.Scan(iRange)
End Sub

'=======================================
' �𓚏����̎��s
' @return True:�z�u����
'=======================================
Private Function Solver_Execute() As Boolean
    Solver_Execute = pInnerObj.Execute()
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
