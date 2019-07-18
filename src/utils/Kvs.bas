Attribute VB_NAME = "Kvs"

' Aggr
' Public NotInherits Class Kvs
' ------------------------------------------------------------------------
'   ���̃��W���[����"config"�Ƃ������O�̃V�[�g������ Key Value Store ��
'   �쐬���܂�
'   config �Ƃ������O�̃V�[�g�ɂ́A���L�̍\���̃e�[�u����5�s�ڂ���n�܂��Ă���
'   ���̂Ƃ݂Ȃ��܂�
'   2��ځF�L�[�ƂȂ镶����
'   3��ځF�l�ƂȂ镶����
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

    ' TODO: IMPLEMENTS
    Public Sub setConfig(key As String, value As String)

    End Sub
    
' End Module
