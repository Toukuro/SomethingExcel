VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public X As Long
Public Y As Long

Public Property Get IntX() As Integer
    IntX = CInt(X And &H7FFF)
End Property

Public Property Get IntY() As Integer
    IntY = CInt(Y And &H7FFF)
End Property

Public Sub SetXY(iX As Long, iY As Long)
    X = iX
    Y = iY
End Sub

Public Sub SetXYbyInt(iX As Integer, iY As Integer)
    X = CLng(iX)
    Y = CLng(iY)
End Sub

