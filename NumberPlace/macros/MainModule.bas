Attribute VB_Name = "MainModule"
'---------------------------------------
' �O���[�o���ϐ�
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
' �𓚎��s
'---------------------------------------
Public Sub Solving()
    Dim SolverCollection As New Collection
    Dim solv As Solver
    Dim numRange As Range
    Dim i As Integer
    
    DisableAutomatic
    
    '�Ֆʂ͈̔�
    Set numRange = ActiveSheet.Range(Cells(2, 2), Cells(10, 10))
    
    '��@�p�^�[�����R���N�V�����ɐݒ�
    SolverCollection.Add New EliminationSolver
    SolverCollection.Add New RoundRobinSolver
    
    '���ׂẲ�@���g���Ė�������
    For i = 1 To SolverCollection.Count
        Set solv = SolverCollection(i)
        Call solv.Scan(numRange)
        If solv.Execute Then
            MsgBox "�𓚊���"
            EnableAutomatic
            Exit Sub
        End If
    Next i
    
    MsgBox "�𓚎��s"
    EnableAutomatic
End Sub

'---------------------------------------
' �Ֆʃ��Z�b�g
'---------------------------------------
Public Sub Clear()
    Dim solv As New Solver
    Dim numRange As Range
    
    Set numRange = ActiveSheet.Range(Cells(2, 2), Cells(10, 10))
    Call solv.Clear(numRange)
    
    EnableAutomatic
End Sub

'---------------------------------------
' �����v�Z�̖�����
'---------------------------------------
Private Sub DisableAutomatic()
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False
End Sub

'---------------------------------------
' �����v�Z�̗L����
'---------------------------------------
Private Sub EnableAutomatic()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
