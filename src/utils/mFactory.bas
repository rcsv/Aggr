Attribute VB_Name = "mFactory"

' Aggr
' Factory Module
' ------------------------------------------------------------------------
' last update: 2019-07-11
' rcsv

Option Explicit

Public Function Init(o As Initializable, ParamArray p()) As Object

    Dim p2() As Variant, i
    ReDim p2(UBound(p))
    For i = 0 To UBound(p)
        If IsObject(p(i)) Then
            Set p2(i) = p(i)
        Else
            Let p2(i) = p(i)
    Next i

    Set Init = o.Init(p2)
End Function