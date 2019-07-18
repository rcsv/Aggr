Attribute VB_NAME = "Kvs"

' Aggr
' Public NotInherits Class Kvs
' ------------------------------------------------------------------------
'   �O������F���̃��W���[���́A�V�[�g���uconfig�v�ɏ����ꂽ�������ƂɁA
'   �}�N�����ł�Key Value Store ���쐬���܂��B�ŏ��ɌĂяo���ꂽ�Ƃ��Ƀn�b�V��
'   �}�b�v�����T�����邽�߁A���ϒT�����Ԃ́A���ڃV�[�g��T�������葁���͂�
'   ���L�̍\���̃e�[�u����5�s�ڂ���n�܂��Ă�����̂Ƃ݂Ȃ�
'   2��ځF�L�[�ƂȂ镶����
'   3��ځF�l�ƂȂ镶����
'   Revision:

Option Explicit

' Public Module Kvs

    ' kvs Core
    Private hashMap As Object

    '
    ' default Sheet object name "config"
    Private Const wsKVS_STORAGE As String = "config"

    ' default table spec
    Private Const colKVS_KEY As Integer = 2   ' �L�[��2���
    Private Const colKVS_VALUE As Integer = 3 ' �l��3���
    Private Const rowKVS_START As Integer = 5 ' 5�s�ڂ���n�܂�

    ' public String getConfig(String key)
    ' ��{�I�ɌĂяo�����\�b�h�͂��ꂾ��
    ' @param key �ݒ荀�ږ�
    ' @return ������
    ' @throw UnknownKeyException �L�[�����Ƃ��Ƃ̃e�[�u���ɂȂ���΁A�G���[���A��܂�
    '
    Public Function getConfig(key As String) As String 
        Dim hm As Object
        Set hm = getInstance()
        getConfig = hm.Item(key)
    End Function

    ' private static Kvs getInstance()
    ' �n�b�V���}�b�v�̕�������������邽�߃C���X�^���X�͈��
    ' @return kvs�Ƃ������̃n�b�V���}�b�v�I�u�W�F�N�g
    '
    Private Function getInstance() As Object
        IF hashMap Is Nothing Then
            Call initHM(hashMap)
        End If
        Set getInstance = hashMap
    End Function

    ' initHM
    ' hashMap �I�u�W�F�N�g�����������邾��
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

    ' public void setConfig(String key, String value)
    ' add map tuple data into Dictionary object and worksheet "config".
    ' @param key String
    ' @param value String
    '
    Public Sub setConfig(key As String, value As String)
        Dim i As Integer: i = rowKVS_START
        If Not hashMap.Exists(key) Then
            hashMap.Add key, value
        Else
            hashMap.Item(key) = value
        End If

        Application.ScreenUpdating = False
        With ThisWorkbook.Worksheets(wsKVS_STORAGE)
            Do While .Cells(i, colKVS_KEY) <> ""
                If .Cells(i, colKVS_KEY) Then
                    .Cells(i, colKVS_VALUE) = value
                    End Sub
                End If
                i = i + 1
            Loop
            .Cells(i, colKVS_KEY) = key
            .Cells(i, colKVS_VALUE) = value
        End With
        Application.ScreenUpdating = True
    End Sub
    
' End Module
