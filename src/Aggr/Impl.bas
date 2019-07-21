Attribute VB_Name = "Impl"

' Aggr
' main logic
' ------------------------------------------------------------------------
'
'
'

Option Explicit

' Public Module Impl

    ' --- reset ----------------------------------------------------------
    Public Sub reset()

        ' obtain flush ignorance sheet list
        Dim shName
        Dim skipNames() As Variant
        skipNames = split(Kvs.getConfig("sheet.remove_ignore"), ",")

        Dim i As Integer, j As Integer, shIgnore

        j = Sheets.Count
        For i = 1 To j
            shName = ThisWorkbook.Worksheets(i).Name

            For Each shIgnore In skipNames
                If shName = shIgnore Then
                    Call Utils.dPrint("SKIP REMOVE: " & shName)
                Else
                    Call Utils.wrapDelete(Worksheets(i))
                    i = i - 1
                    j = j - 1
                End IF
            Next shIgnore
        Next i

        Call Utils.dPrint("flush data")
        
        Application.ScreenUpdating = False
        ThisWorkbook.worksheet("saturn") ' . Clear

    End Sub

    ' --- import ---------------------------------------------------------
    Public Sub import()

    End Sub

    ' --- Sync -----------------------------------------------------------
    Public Sub sync()

    End Sub


' End Module