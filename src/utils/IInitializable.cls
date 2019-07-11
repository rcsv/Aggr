VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "IInitializable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Aggr
' public interface Initializable
' ------------------------------------------------------------------------
' last update: 2019-07-11
' rcsv

Option Explicit

' Public Class MustInherit Initializable

    Public Function Init(p() As Variant) As Object
    End Function

' End Class