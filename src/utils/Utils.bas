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
        ptJpBrackets = 5 ' 【】日本語のフォントで出るブラケット
        ptJpBraces = 6   ' 「」日本語の鍵括弧
        ptAngleBrackets = 7 ' <> 大小記号で構成するブラケット
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
                strSTART = "【": strEND = "】"
            Case ptParen
                strSTART = "(": strEND = ")"
            Case ptBrace
                strSTART = "{": strEND = "}"
            Case ptBracket
                strSTART = "[": strEND = "]"
            Case ptJpBrace
                strSTART = "「": strEND = "」"
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
    ' 例えばファイル名の末尾に確認済みのマークを追加するなどに使用する
    ' 既に確認マークが同じ場所に付与されている場合は文字列の変更をしない
    ' @param str 確認済みマークを挿入したい文字列
    ' @param flgChar 確認済みマークとする文字（デフォルトは黒丸●）
    ' @retun 確認済みマーク付与済みの文字列。すでに付与されている場合は変更なし
    Public Function addFLAG(str As String, _
            Optional flgChar As String = "●") As String

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
    ' ワークシートを直接操作して、任意の範囲のテキストフォーマットや条件付き書式をリセットする
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
    ' マクロから別のエクセルファイルを開く
    ' @param 開くエクセルファイルへのフルパス
    ' @return 開いた後のエクセル workbook オブジェクト
    Public Function wrapOpen(target As String) As Workbook
        dPrint "try to open: " & target
        Application.DisplayAlerts = False
        Set wrapOpen = Workbooks.Open(target, False)
        Application.DisplayAlerts = True
    End Function

    ' public static void wrapSave(Workbook target)
    ' マクロから保存する
    Public Sub wrapSave(target As Workbook)
        Application.DisplayAlerts = False
        target.Save
        Application.DisplayAlerts = True
    End Sub

    ' public static void wrapClose(Workbook target)
    ' ワークブックオブジェクトからファイルを閉じる
    ' @param Workbookオブジェクト
    Public Sub wrapClose(target As Workbook)
        Application.DisplayAlerts = False
        target.Close
        Application.DisplayAlerts = True
    End Sub



' End Module