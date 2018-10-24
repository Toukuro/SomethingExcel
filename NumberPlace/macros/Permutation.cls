VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Permutation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------
' 定数
'---------------------------------------
Private Const RNO_COUNT As Long = 1
Private Const RNO_DELCNT As Long = 2
Private Const RNO_DATATOP As Long = 3

'---------------------------------------
' メンバー変数
'---------------------------------------
Private workRange As Range

'---------------------------------------
' 初期化
'---------------------------------------
Private Sub Class_Initialize()
    Set workRange = Sheets("Work").Cells
    Call workRange.Range("A:I").ClearContents
End Sub

'---------------------------------------
' 格納データのクリア
'---------------------------------------
Public Sub Clear()
    Dim i As Integer
    
    For i = 1 To 9
        Call SetCount(i, 0)
    Next i
End Sub

'---------------------------------------
' 順列パターンの生成
'---------------------------------------
Public Sub GenPattern(iIdx As Integer, iSeed As String, Optional iInterim As String = "")
    Dim nextInterim As String
    Dim leftSeed As String
    Dim seedLen As Integer
    
    seedLen = Len(iSeed)
    If seedLen >= 9 Then
        workRange.Range("J1:J362882").Copy workRange(RNO_COUNT, iIdx)
        Exit Sub
    ElseIf seedLen <= 0 Then
        Call AddPattern(iIdx, iInterim)
        Exit Sub
    End If

    For i = 1 To seedLen
        nextInterim = iInterim & mid(iSeed, i, 1)
        leftSeed = left(iSeed, i - 1) & mid(iSeed, i + 1)

        Call GenPattern(iIdx, leftSeed, nextInterim)
    Next i
End Sub

'---------------------------------------
' 順列パターンの追加
'---------------------------------------
Public Sub AddPattern(iIdx As Integer, iPattern As String)
    Dim cnt As Long
    
    cnt = Count(iIdx) + 1
    workRange(RNO_DATATOP + cnt - 1, iIdx).Value = iPattern
    Call SetCount(iIdx, cnt)
End Sub

'---------------------------------------
' パターンデータの削除
'---------------------------------------
Public Sub DelPattern(iIdx As Integer, iPNo As Long)
    Dim cnt As Long
    
    workRange(RNO_DATATOP + iPNo - 1, iIdx).Value = ""
    cnt = DelCount(iIdx) + 1
    Call SetDelCount(iIdx, cnt)
    
    DoEvents
End Sub

'---------------------------------------
' パターンデータ削除した行の削除
'---------------------------------------
Public Sub Shrink(iIdx As Integer)
    Dim allCnt As Long, delCnt As Long, tgtCnt As Long
    Dim delStart As Long, delEnd As Long
    Dim delAddr As String
    Dim pNo As Long
    
    delCnt = DelCount(iIdx)
    If delCnt <= 0 Then
        Exit Sub
    End If
    
    allCnt = Count(iIdx)
    delEnd = 0
    For pNo = (RNO_DATATOP + allCnt - 1) To RNO_DATATOP Step -1
        If delEnd > 0 Then
            '削除範囲設定中
            If workRange(pNo, iIdx).Value <> "" Then
                delStart = pNo + 1
                                
                delAddr = CellAddr(iIdx, delStart, iIdx, delEnd)
                Call workRange.Range(delAddr).Delete(XlDeleteShiftDirection.xlShiftUp)
                
                '削除したセル数
                tgtCnt = delEnd - delStart + 1
                delCnt = delCnt - tgtCnt
                allCnt = allCnt - tgtCnt
                
                Call SetDelCount(iIdx, delCnt)
                Call SetCount(iIdx, allCnt)
                DoEvents
                
                delEnd = 0
            End If
        Else
            '削除範囲設定前
            If workRange(pNo, iIdx).Value = "" Then
                delEnd = pNo
            End If
        End If
    Next pNo

    If delEnd > 0 Then
        '削除範囲設定中
        delStart = RNO_DATATOP
        
        delAddr = CellAddr(iIdx, delStart, iIdx, delEnd)
        Call workRange.Range(delAddr).Delete(XlDeleteShiftDirection.xlShiftUp)
        
        '削除したセル数
        tgtCnt = delEnd - delStart + 1

        Call SetDelCount(iIdx, DelCount(iIdx) - tgtCnt)
        Call SetCount(iIdx, Count(iIdx) - tgtCnt)
        DoEvents
    End If
End Sub

'---------------------------------------
' 順列パターンの取得
'---------------------------------------
Public Function GetPattern(iIdx As Integer, iPNo As Long) As String
        GetPattern = workRange(RNO_DATATOP + iPNo - 1, iIdx).Value
End Function

'---------------------------------------
' パターン数の取得／設定
'---------------------------------------
Public Function Count(iIdx As Integer) As Long
    Count = workRange(RNO_COUNT, iIdx).Value
End Function

Private Sub SetCount(iIdx As Integer, iValue As Long)
    workRange(RNO_COUNT, iIdx).Value = iValue
End Sub

'---------------------------------------
' 削除パターン数の取得／設定
'---------------------------------------
Public Function DelCount(iIdx As Integer) As Long
    DelCount = workRange(RNO_DELCNT, iIdx).Value
End Function

Private Sub SetDelCount(iIdx As Integer, iValue As Long)
    workRange(RNO_DELCNT, iIdx).Value = iValue
End Sub

'---------------------------------------
' C1R1形式のセルアドレス式を取得
'---------------------------------------
Private Function CellAddr(iCol1 As Integer, iRow1 As Long, Optional iCol2 As Integer = 0, Optional iRow2 As Long = 0) As String
    CellAddr = mid("ABCDEFGHI", iCol1, 1) & CStr(iRow1)
    If iCol2 > 0 And iRow2 > 0 Then
        CellAddr = CellAddr & ":" & mid("ABCDEFGHI", iCol2, 1) & CStr(iRow2)
    End If
End Function
