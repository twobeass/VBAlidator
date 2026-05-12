Attribute VB_Name = "DeepUdtChain"
Option Explicit

Private Type Lvl4
    n As Long
End Type

Private Type Lvl3
    a As Lvl4
End Type

Private Type Lvl2
    b As Lvl3
End Type

Private Type Lvl1
    c As Lvl2
End Type

Sub DeepLeaf()
    Dim x As Lvl1
    Dim n As Long
    n = x.c.b.a.bogus   ' bogus is not a member of Lvl4 — depth-5 chain must still type-check
End Sub
