Attribute VB_NAME = "EntryPoint"

' Aggr
' EntryPoint
' ------------------------------------------------------------------------
'   aggregate same formatting files into one excels
'

Option Explicit

' Public Module EntryPoint

    ' records
    Private arrRecords() As Variant

    ' Controller Mode
    Public Enum AppFormMode
        import_files = 1 ' filename
        generate_id = 2 ' theme id generate and distribute back
        sync = 3
    End Enum

    ' EntryPoints
    ' --------------------------------------------------------------------
    '

    ' EntryPoint 1: import files
    ' just importing, not add information
    Public Sub ep_import()
        If MsgBox(Kvs.getConfig("msg.import.start"), vbOkCancel + vbQuestion) = vbCancel Then
            Exit Sub
        End If

        Call ep(import_files)
    End Sub

    ' EntryPoint 2: add ID
    ' publish unique id to each records with imporing files
    Public Sub ep_ID()
        If MsgBox(Kvs.getConfig("msg.add_id.start"), vbOkCancel + vbQuestion) = vbCancel Then
            Exit Sub
        End If

        Call ep(generate_id) 
    End Sub

    ' EntryPoint 3: sync
    ' sync all data
    Public Sub ep_sync()
        If MsgBox(Kvs.getConfig("msg.sync.start"), vbOkCancel + vbQuestion) = vbCancel Then
            Exit Sub
        End If

        Call ep(sync)
    End Sub

    ' real entry point
    ' --------------------------------------------------------------------
    '

    ' EntryPoint 4: general
    ' a facade pattern of entrypoint
    Private Sub ep(ByVal switch As AppFormMode)
        Err.Clear
        On Error GoTo ERR_HANDLE

        Call impl.reset
        Call impl.import

        Select Case switch
            Case import_files
                Call impl.stat
            Case generate_id
                Call impl.generateId
            Case sync
                Call impl.sync
        End Select

        Exit Sub
    ERR_HANDLE:
        Call Utils.dPrint("err# " & Err.Number & vbTab & ", Description: " & Err.Description)
        MsgBox "Error: " & Err.Number & Chr(13) & Err.Description
        Application.ScreenUpdating = True
    End Sub

' End Module