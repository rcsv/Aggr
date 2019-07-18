Attribute VB_NAME = "Kvs"

' Aggr
' Public NotInherits Class Kvs
' ------------------------------------------------------------------------
'   このモジュールは"config"という名前のシートを見て Key Value Store を
'   作成します
'   config という名前のシートには、下記の構成のテーブルが5行目から始まっている
'   ものとみなします
'   2列目：キーとなる文字列
'   3列目：値となる文字列
'   Key-Value-Store Implementation
'   Revision:

Option Explicit

' Public Module Kvs

    ' kvs Core
    Private hashMap As Object

    '
    ' default Sheet object name "config"
    Private Const wsKVS_STORAGE As String = "config"

    ' default table spec
    Private Const colKVS_KEY As Integer = 2   ' キーは2列目
    Private Const colKVS_VALUE As Integer = 3 ' 値は3列目
    Private Const rowKVS_START As Integer = 5 ' 5行目から始まる

    ' public String getConfig(String key)
    ' 基本的に呼び出すメソッドはこれだけ
    ' @param key 設定項目名
    ' @return 文字列
    ' @throw UnknownKeyException キーがもともとのテーブルになければ、エラーが帰ります
    '
    Public Function getConfig(key As String) As String 
        Dim hm As Object
        Set hm = getInstance()
        getConfig = hm.Item(key)
    End Function

    ' private static Kvs getInstance()
    ' ハッシュマップの複数生成を避けるためインスタンスは一つ
    ' @return kvsという名のハッシュマップオブジェクト
    '
    Private Function getInstance() As Object
        IF hashMap Is Nothing Then
            Call initHM(hashMap)
        End If
        Set getInstance = hashMap
    End Function

    ' initHM
    ' hashMap オブジェクトを初期化するだけ
    Private Sub initHM(ByRef hm As Object)
        Set hm = CreateObject("Scripting.Dictionary")
        Dim i, key As String, value As String
        i = rowKVS_START
        key = "" : value = ""

        With ThisWorkbook.Worksheets(wsKVS_STORAGE)
            Do While .Cells(i, colKVS_KEY) <> ""
                key = .Cells(i, colKVS_KEY)
                value = .Cells(i, colKVS_VALUE)
                hm.Add key, value
                i = i + 1
            Loop
        End With
    End Sub

    ' TODO: IMPLEMENTS
    Public Sub setConfig(key As String, value As String)

    End Sub
    
' End Module
