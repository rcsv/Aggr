Attribute VB_NAME = "Utils"

' Aggr
' Public NotInherits Class Utils
' ------------------------------------------------------------------------
'   last update: 2019-07-11
'

Option Explicit

' Public Module Utils

    ' default log filename is:
    private Const strLOGFILE As String = "\transaction.log"

    Public Enum pType
        ptParen = 1      ' () parentheses
        ptBrace = 2      ' {} braces
        ptBracket = 3    ' [] brackets
        ptQuote = 4      ' '' single quotes
        ptJpBrackets = 5 ' �y�z���{��̃t�H���g�ŏo��u���P�b�g
        ptJpBraces = 6   ' �u�v���{��̌�����
        ptAngleBrackets = 7 ' <> �召�L���ō\������u���P�b�g
    End Enum

    ' public static void dPrint(String msg)
    ' write a message down to both the debug console and a log file
    '
    Public Sub dPrint(msg As String)
        Dim buf, path, objFSO As Object

        buf = Now & vbTab & msg
        Debug.Print buf
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        path = ThisWorkbook.path & strLOGFILE

        ' It probably fails when path include WebDAV-related directory
        With objFSO
            If Not .FileExists(path) Then .CreateTextFile(path)
            With .OpenTextFile(path, 8)
                .WriteLine buf
                .Close
            End With
        End With

        Set objFSO = Nothing
    End Sub

    ' public static String getBraceInside(String str)
    ' distilling string nested braces
    ' @param String string within pair of brackets
    ' @param pType (Optional) parenthesis type
    ' @return string within the parentheses
    '
    Public Function getBraceInside(ByVal str As String, _
            Optional paren_type As pType = jp_bracket) As String
    
        Dim pos_start As Integer, pos_end As Integer
        Dim strSTART As String, strEND As String
        getBracketInside = ""

        Select Case paren_type
            Case ptJpBracket
                strSTART = "�y": strEND = "�z"
            Case ptParen
                strSTART = "(": strEND = ")"
            Case ptBrace
                strSTART = "{": strEND = "}"
            Case ptBracket
                strSTART = "[": strEND = "]"
            Case ptJpBrace
                strSTART = "�u": strEND = "�v"
            Case ptAngleBracket
                strSTART = "<": strEND = ">"
            Case ptQuote
                strSTART = "'": strEND = "'"
            Case Else
                ' do nothing?
                Err.Raise 515, TypeName(Me), "Unknown Parenthesis Type:" & paren_type
        End Select

        pos_start = InStr(str, strSTART)
        pos_end = InStr(str, strEND)
        pos_end = pos_end - pos_start - 1

        getBraceInside = Mid(str, (pos_start +1), pos_end)
        Call dPrint("String " & getBraceInside & " distilled from: " & str)
    End Function

    ' public static String addFLAG(String str)
    ' �Ⴆ�΃t�@�C�����̖����Ɋm�F�ς݂̃}�[�N��ǉ�����ȂǂɎg�p����
    ' ���Ɋm�F�}�[�N�������ꏊ�ɕt�^����Ă���ꍇ�͕�����̕ύX�����Ȃ�
    ' @param str �m�F�ς݃}�[�N��}��������������
    ' @param flgChar �m�F�ς݃}�[�N�Ƃ��镶���i�f�t�H���g�͍��ہ��j
    ' @retun �m�F�ς݃}�[�N�t�^�ς݂̕�����B���łɕt�^����Ă���ꍇ�͕ύX�Ȃ�
    Public Function addFLAG(str As String, _
            Optional flgChar As String = "��") As String

        Dim posLastDot, strExt
        posLastDot = InStrRev(str, ".")
        strExt = Right(str, Len(str) - posLastDot)

        If Mid(str, posLastDot - 1, 1) <> flgChar Then
            addFLAG = Left(str, posLastDot - 1) & flgChar & "." & strExt
        Else
            addFLAG = str
        End If
    End Function

    ' public static void ResetTextFormatting(Worksheet ws)
    ' ���[�N�V�[�g�𒼐ڑ��삵�āA�C�ӂ͈̔͂̃e�L�X�g�t�H�[�}�b�g������t�����������Z�b�g����
    ' @param ws a worksheet object
    Public Sub resetTextFormatting(ws As Worksheet)
        Dim tempA As Integer, tempB As Integer
        Dim r1 As Range, r2 As Range

        tempA = getConfig("column.start")
        tempB = getConfig("column.end")

        Set r1 = ws.Cells(tempA, 1)
        Set r2 = ws.Cells(50, tempB)    ' TODO: getEND OF X
        ws.Range(r1, r2).Select
        Cells.FormatConditions.Delete
    End Sub

    ' public static Workbook wrapOpen(String filename)
    ' �}�N������ʂ̃G�N�Z���t�@�C�����J��
    ' @param �J���G�N�Z���t�@�C���ւ̃t���p�X
    ' @return �J������̃G�N�Z�� workbook �I�u�W�F�N�g
    Public Function wrapOpen(target As String) As Workbook
        dPrint "try to open: " & target
        Application.DisplayAlerts = False
        Set wrapOpen = Workbooks.Open(target, False)
        Application.DisplayAlerts = True
    End Function

    ' public static void wrapSave(Workbook target)
    ' �}�N������ۑ�����
    Public Sub wrapSave(target As Workbook)
        Application.DisplayAlerts = False
        target.Save
        Application.DisplayAlerts = True
    End Sub

    ' public static void wrapClose(Workbook target)
    ' ���[�N�u�b�N�I�u�W�F�N�g����t�@�C�������
    ' @param Workbook�I�u�W�F�N�g
    Public Sub wrapClose(target As Workbook)
        Application.DisplayAlerts = False
        target.Close
        Application.DisplayAlerts = True
    End Sub



' End Module