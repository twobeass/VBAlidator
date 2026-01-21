Attribute VB_Name = "CWireService"
Option Explicit

Public Sub UpdateWire(vsoShape As Visio.shape)
    ' ...
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
    Dim formulaStr As String
    If isBegin Then formulaStr = cable.Cells("BeginX").FormulaU Else formulaStr = cable.Cells("EndX").FormulaU
    IsEndConnected = (InStr(formulaStr, "PAR(PNT(") > 0)
End Function

Public Function IsCableConnected(ByVal connector As Visio.shape) As Boolean
    On Error GoTo ErrHandler
    Dim fromConnected As Boolean
    Dim toConnected As Boolean

    ' Check if the beginning of the connector is connected
    fromConnected = connector.Connects.count > 0 And connector.Cells("BeginX").FormulaU Like "PAR(PNT(*"
    ' Check if the end of the connector is connected
    toConnected = connector.Connects.count > 0 And connector.Cells("EndX").FormulaU Like "PAR(PNT(*"

    ' Return True if both ends are connected
    IsCableConnected = fromConnected And toConnected
    Exit Function

ErrHandler:
    Call Lib_ErrorHandler.HandleError("CWireService.IsCableConnected", Err.Description)
End Function

Public Function GetUnconnectedCables(ByVal page As Visio.page) As Collection
    On Error GoTo ErrHandler
    Dim shp As Visio.shape
    Dim col As New Collection
    Dim isConnected As Boolean

    If page Is Nothing Then Set page = Visio.activePage

    ' Iterate through all shapes on the page
    For Each shp In page.shapes
        If Not shp.master Is Nothing Then
            If shp.master.NameU Like "Cable*" Then
                If Not shp.master.NameU Like "Cable_OPR*" Then
                     isConnected = IsCableConnected(shp)
                     If Not isConnected Then
                         col.Add shp
                     End If
                End If
            End If
        End If
    Next shp

    Set GetUnconnectedCables = col
    Exit Function

ErrHandler:
    Call Lib_ErrorHandler.HandleError("CWireService.GetUnconnectedCables", Err.Description)
End Function
