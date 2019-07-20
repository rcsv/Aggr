VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 ' True
END
Attribute VB_NAME = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Aggr
' Record
' ------------------------------------------------------------------------
'   Record represents a dataset of table
'

Option Explicit

' Public Class Record
Implements IInitializable

    ' record title
    Private mTitle As String

    ' crues for reaching to original files
    Public dist As String

    ' category 1
    Public mC1 As ICategory

    ' category 2
    Public mC2 As ICategory

    ' configuration key name for number of Column LEFT-END
    ' prefer 1, not zero
    Private Const keyCOL_LEFT As String = "column.start" 
    Private colLEFT As Integer
    
    ' configuration key name for number of Column RIGHT-END
    ' probably 120
    Private Const keyCOL_RIGHT As String = "column.end" 
    Private colRIGHT As Integer

    ' configuration key name for number of Row table starts
    Private Const keyROW_START As String = "row.start"
    Private rowSTART As Integer
    
    ' TITLE cannot set from outer zone
    Private Property Let Title(key As String)
        mTitle = key
    End Property

    Public Property Get Title()
        Title = mTitle
    End Property

    ' check flag whether this record is enable/disable for some reason
    Private Property Let IDisabled_flgDisabled(ByVal flag As Boolean)
        IDisabled_flgDisabled = flag
    End Property

    ' Init
    ' @return this
    Function Init(p() As Variant) As Object
        Init = IInitializable_Init(p)
    End Function

    ' IInitializable_Init
    ' initialize
    ' @param p Array of arguments
    ' @return this
    Private Function IInitializable_Init(p() As Variant) As Object

    End Function

' End Class
