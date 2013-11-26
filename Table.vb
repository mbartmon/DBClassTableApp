Public Class Table

    Private Name As String
    Private Cols() As columns = Nothing

    Public Sub New(ByVal name As String, ByVal cn As OleDb.OleDbConnection)

        Dim cmd As New OleDb.OleDbCommand
        Dim reader As OleDb.OleDbDataReader
        Dim dt As DataTable
        Try
            cmd.Connection = cn
            cmd.CommandText = "SELECT* FROM [" & name & "]"
            reader = cmd.ExecuteReader(CommandBehavior.KeyInfo)
            dt = reader.GetSchemaTable
        Catch
            Exit Sub
        End Try

        Me.Name = name
        Dim c As Integer = 0
        Dim S As String = Nothing
        For Each dr As DataRow In dt.Rows
            ReDim Preserve Cols(c)
            Cols(c) = New columns(PROPERTY_PREFIX & CType(dr("ColumnName"), System.String), dr("DataType").ToString, CType(dr("IsKey"), System.Boolean), CType(dr("isAutoIncrement"), System.Boolean), CType(dr("ColumnSize"), System.Int64))
            c += 1
        Next

    End Sub

    Public Sub New(name As String, cn As SqlClient.SqlConnection)

        Dim xsql As String = "SELECT c.name 'ColumnName', t.Name 'DataType', c.max_length 'MaxLength', c.precision, c.scale, c.is_nullable," & _
                            " ISNULL(i.is_primary_key, 0) 'PrimaryKey' FROM sys.columns c INNER JOIN sys.types t ON c.system_type_id = t.system_type_id" & _
                            " LEFT OUTER JOIN sys.index_columns ic ON ic.object_id = c.object_id AND ic.column_id = c.column_id" & _
                            " LEFT OUTER JOIN sys.indexes i ON ic.object_id = i.object_id AND ic.index_id = i.index_id" & _
                            " WHERE c.object_id = OBJECT_ID('" & name & "')"
        Dim da As New SqlClient.SqlDataAdapter(xsql, sqlCN)
        Dim ds As New DataSet(name)
        Dim dict As New Dictionary(Of String, String)
        Try
            da.Fill(ds, name)
        Catch
            Exit Sub
        End Try

        Me.Name = name
        Dim c As Integer = 0
        Dim S As String = Nothing
        For Each dr As DataRow In ds.Tables(0).Rows
            If Not dict.ContainsKey(dr("ColumnName")) And dr("ColumnName") <> "upsize_ts" Then
                ReDim Preserve Cols(c)
                Cols(c) = New columns(PROPERTY_PREFIX & CType(dr("ColumnName"), System.String), dr("DataType").ToString, CType(dr("PrimaryKey"), System.Boolean), CType(dr("PrimaryKey"), System.Boolean), CType(dr("MaxLength"), System.Int64))
                c += 1
                dict.Add(dr("ColumnName"), dr("ColumnName"))
            End If
        Next

    End Sub
    Public Function getName() As String
        Return Me.Name
    End Function
    Public Function getCols() As columns()
        Return Cols
    End Function
    Public Function getCol(ByVal index As Integer) As columns
        Return Cols(index)
    End Function

End Class

Public Class columns
    Private Name As String
    Private dbType As String
    Private isKey As Boolean = False
    Private isAutoNumber As Boolean = False
    Private columnSize As Long = 0

    Public Sub New()

    End Sub

    Protected Friend Sub New(ByVal name As String, ByVal type As String, ByVal isKey As Boolean, ByVal isAutoNumber As Boolean, columnSize As Long)
        Me.Name = name
        Me.dbType = type
        Me.isKey = isKey
        Me.isAutoNumber = isAutoNumber
        Me.columnSize = columnSize
    End Sub

    Public Sub setName(ByVal name As String)
        Me.Name = name
    End Sub
    Public Function getName() As String
        Return Me.Name
    End Function
    Public Sub setDbType(ByVal type As String)
        Me.dbType = type
    End Sub
    Public Function getDbType() As String
        Return Me.dbType
    End Function
    Public Sub setIsKey(ByVal isKey As Boolean)
        Me.isKey = isKey
    End Sub
    Public Function getIsKey() As Boolean
        Return Me.isKey
    End Function
    Public Sub setIsAutoNumber(ByVal isAutoNumber As Boolean)
        Me.isAutoNumber = isAutoNumber
    End Sub
    Public Function getIsAutoNumber() As Boolean
        Return Me.isAutoNumber
    End Function
    Public Sub setColumnSize(value As Long)
        Me.columnSize = value
    End Sub
    Public Function getColumnSize() As Long
        Return Me.columnSize
    End Function

End Class

