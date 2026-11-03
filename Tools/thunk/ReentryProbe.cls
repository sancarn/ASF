VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ThunkReentryProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public engine As ASF
Public inner As Long
Public calls As Long
Public releases As Long
Public Property Get Value() As Double
Attribute Value.VB_UserMemId = 0
    calls = calls + 1
    engine.Run inner
    Value = CDbl(engine.OUTPUT_)
End Property
Public Function Make() As Object
    Dim item As New ThunkLifetime
    Set item.parent = Me
    Set Make = item
End Function
Public Sub Released()
    releases = releases + 1
    engine.Run inner
End Sub
