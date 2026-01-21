Attribute VB_Name = "CWireService"
Option Explicit

Public Sub UpdateWire(vsoShape As Visio.shape)
    ' ... trimmed ...
End Sub

Private Function TryAutoConnectCableEnds(cable As Visio.shape, pg As Visio.page) As Boolean
    On Error GoTo ErrHandler
    Dim beginConnected As Boolean, endConnected As Boolean
    Dim madeConnection As Boolean

    beginConnected = IsEndConnected(cable, True)
    endConnected = IsEndConnected(cable, False)

    TryAutoConnectCableEnds = madeConnection
    Exit Function
ErrHandler:
    TryAutoConnectCableEnds = False
End Function

Private Function IsEndConnected(cable As Visio.shape, isBegin As Boolean) As Boolean
    On Error Resume Next
    IsEndConnected = True
End Function
