
' Aggr
' Imports source code automatically
' ------------------------------------------------------------------------
'   i18n   
' HOW TO USE
' insert this source code into "ThisWorkbook"
'
'

Option Explicit

' Public Class ThisWorkbook

    ' Depencency Injection file
    Private Const CONF_SRC As String = "deployment.txt"

    ' Filename of Export ThisWorkbook
    Private Const EXPORT_ME As String = "ThisWorkbook-sjis-cls"

    ' flag for enabling "ThisWorkbook_Open" Event 
    Private Const ENABLE_EVENT_OPEN As Boolean = True

    ' Shortcut Key
    Private Const SHORTCUTKEY_RELOAD As String = "r"

    ' Event Triggers
    ' --------------------------------------------------------------------
    '

    ' Workbook Open
    Private Sub Workbook_Open()
        If ENABLE_EVENT_OPEN = False Then
            Exit Sub
        End If

        Call setShortKey
        Call reloadModule
    End Sub

    ' WorkBook Before Close
    Private Sub Workbook_BeforeClose(Cancel As Boolean)
        Call clearShortKey
    End Sub

    ' Prepared Methods
    ' --------------------------------------------------------------------
    '

    ' manual reload
    Public Sub reloadModule()
        Dim msgError = loadModule("." & Application.PathSeparator & CONF_SRC)

        If Len(msgError) > 0 Then
            MsgBok msgError
        End If
    End Sub

    ' manual export me
    Public Sub exportThisWorkbook
        Call exportModule("ThisWorkbook", SHORTCUTKEY_RELOAD) 
    End Sub

    ' main procedures
    ' --------------------------------------------------------------------
    '

    ' loadModule
    ' read configuration file for deployment and load all files into xlsm file
    Private Function loadModule(ByVal pathConf As String) As String
        Dim isClear As Boolean
        isClear = clearModule

        If isClear = False Then
            loadModule = "Error: Fail to Remove all modules."
            Exit Function
        End If

        ' check whether module file list is
        pathConf = absPath(pathConf)

        Dim isExistList As Boolean
        isExistList = checkExistFile(pathConf)

        If isExistList = False Then
            loadModule = "Error: missing DI file in: " & pathConf
            Exit Function
        End If

        Dim arrayModules As Variant
        arrayModules = list2array(pathConf)

        If UBound(arrayModules) = 0 Then
            loadModule = "Error: missing file description in the confiig file"
            Exit Function
        End If

        ' loading
        Dim i As Integer
        Dim msgError = ""
        
        Dim strPath As Variant
        For Each strPath In arrayModules
            strPath = absPath(strPath)
            If checkExistFile(strPath) Then
                ThisWorkbook.VBProject.VBComponents.Import strPath
            Else
                msgError = msgError & strPath & " is missing." & vbCrLf
            End If
        Next

        loadModule = msgError

    End Function

    Private Sub exportModule(ByVal nameModule As String, ByVal nameFile As String)
        Dim component As Object
        For Each component In ThisWorkbook.VBProject.VBComponents

            If component.Name = nameModule Then
                component.Export ThisWorkbook.path & _ 
                        Application.PathSeparator & _
                        nameFile
                ' export done.
            End If
        Next
    End Sub

    ' clearFunction
    ' remove all loaded standard/class modules of xlsm
    ' @return boolean success or fail
    Private Function clearModules() As Boolean
        Dim cmp As Object

        ' remove all standard (Type=1), and class (Type=2) Modules
        For Each cmp In ThisWorkbook.VBProject.VBComponents
            If cmp.Type = 1 Or cmp.Type = 2 Then
                ThisWorkbook.VBProject.VBComponents.Remove cmp
            End If
        Next cmp

        Dim nBas As Long
        nBas = countBases()

        Dim nClass As Long
        nClass = countClasses()

        If nBas = 0 And nClass = 0 Then
            clearModules = True
        Else
            clearModules = False
        End If

    End Function

    ' countBases (shorthand)
    ' count standard modules still exist inside
    private Function countBases() As Long
        Dim cnt As Long: cnt = countComponents(1) ' bas=1
        countBases = cnt
    End Function

    ' countClasses (shorthand)
    ' count class modules still exist inside
    Private Function countClasses() As Long
        Dim cnt As Long: cnt = countComponents(2) ' class
        countClasses = cnt
    End Function

    ' countComponents
    ' count specified modules
    Private Function countComponents(ByVal numtype As Integer) As Long
        Dim i As Long
        Dim count As Long
        count = 0

        With ThisWorkbook.VBProject
            For i = 1 To .VBComponents.count
                If .VBComponents(i).Type = numType Then
                    count = count + 1
                End If
            Next
        End With

        countComponents = count
    End Function

    ' convert file path from related to absolute
    Private Function absPath(ByVal pathFile As String) As String
        Dim os As String
        os = Application.OperatingSystem

        ' path
        pathFile = Replace(pathFile, Chr(92), Application.PathSeparator)
        pathFile = Replace(pathFile, "/", Application.PathSeparator)

        Select Case Left(pathFile, 1)
            Case ".":
                Select Case Left(pathFile, 2)
                    Case "..":
                        absPath = ThisWorkbook.path & Application.PathSeparator & pathFile
                    Case Else:
                        absPath = ThisWorkbook.path & Mid(pathFile, 2, Len(pathFile) - 1)
                        Exit Function

                End Select

            Case Application.PathSeparator:
                absPath = pathFile
                Exit Function
        End Select

        If os Like "Windows * " And Left(pathFile, 2) Like "[A-z]" & Application.PathSeparator Then
            absPath = Replace(pathFile, Application.PathSeparator, ":", 1, 1)
            Exit Function
        End If

        If Left(pathFile, 1) Like "[0-9]" Or Left(pathFile, 1) Like "[A-z]" Then
            absPath = ThisWorkbook.path & Application.PathSeparator & pathFile
        Else
            debug.print TypeName(Me) & " fail to get absolute path."
        End If
    End Function

    ' check whether file exists here
    Private Function checkExistFile(ByVal pathFile As String) As Boolean
        On Error GoTo ERR_DIR
        If Dir(pathFile) = "" Then
            checkExistFile = False
        Else
            checkExistFile = True
        End If

ERR_DIR:
        checkExistFile = False
    End Function

    ' list 2 array
    Private Function list2array(ByVal pathFile As String) As Variant
        Dim os As String
        os = Application.OperatingSystem

        Dim fp As Integer
        fp = FreeFile
        open pathFile For Input As #fp

        Dim arrayOutput() As String
        Dim countLine As Integer
        countLine = 0

        ReDim Preserve arrayOutput(countLine)

        Do Until EOF(fp)
            Dim strLine As String
            Line Input #fp, strLine

            Dim isLf As Long
            is Lf = InStr(strLine, vbLf)

            If os Like "Windows *" And Not isLf = 0 Then
            Else
                strLine = Replace(strLine, vbCr, "")

                If Not Left(strLine, 1) = "'" And Len(strLine) > 0 Then
                    countLine = countLine + 1
                    ReDim Preserve arrayOutput(countLine)
                    arrayOutput(countLine - 1) = strLine
                End If
            End If
        Loop

        Close #fp

        list2array = arrayOutput
    End Function

    ' Short Cut Key Setting
    ' --------------------------------------------------------------------
    '

    ' set shortcut
    Private Sub setShortKey()
        If Application.OperatingSystem Like "Windows *" Then
            Application.MacroOptions _
                    Macro:="ThisWorkbook.reloadModule", _
                    ShortcutKey:=SHORTKEY_RELOAD
        Else
            Application.OnKey _
                    "^" & SHORTKEY_RELOAD, _
                    "ThisWorkbook.reloadModule"

        End If
    End Sub

    ' clear shortcut
    Private Sub clearShortKey()
        If Application.OperationSystem Like "Windows *" Then
            Application.MacroOptions _
                    Macro:="ThisWorkbook.reloadModule", _
                    ShortcutKey:=""
        Else
            Application.OnKey _
                    SHORTKEY_RELOAD, ""
        End If
    End Sub

' End Class