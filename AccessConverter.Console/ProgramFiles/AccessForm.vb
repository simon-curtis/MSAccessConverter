Option Strict Off

Public Class AccessForm
    Inherits Form
    Public ReadOnly Property CurrentRecord As Long = 0
    Public ReadOnly Property RecordSource As String = ""
    Public ReadOnly Property RowFilter As String = ""
    Public Property Recordset As ADODB.Recordset
    Public ReadOnly Property RecordsetType As Access.Dao.RecordsetTypeEnum = Access.Dao.RecordsetTypeEnum.dbOpenDynaset
    Public Property WhereCondition As String = Nothing

    Public Sub New()
    End Sub

    Public Sub New(rowFilter As String)
        Me.RowFilter = rowFilter
    End Sub

    Public Sub NewRecord()

    End Sub
End Class
