Public Class clsDB

    Protected Const DS_DATA_SET = "LeaseLetters"
    Protected loadStmt As String
    Protected loadAllStmt As String
    Protected insertStmt As String
    Protected updateStmt As String
    Protected localTable As String
    Protected priKeyName As String
    Protected created As Boolean

    Protected SQL_DELETE As String = "DELETE FROM " & _
                                    localTable & _
                                    " WHERE " & priKeyName & _
                                    " = ?;"

    Public Sub New(ByVal table As String, ByVal priKeyName As String, ByVal loadStmt As String, ByVal loadAllStmt As String, ByVal insertStmt As String, ByVal updateStmt As String)

        Me.localTable = table
        Me.priKeyName = priKeyName
        Me.loadStmt = loadStmt
        Me.loadAllStmt = loadAllStmt
        Me.insertStmt = insertStmt
        Me.updateStmt = updateStmt
        SQL_DELETE = "DELETE FROM " & _
                     localTable & _
                     " WHERE " & priKeyName & _
                     " = ?;"
    End Sub

    Protected Shared Function prepareStatement(ByVal sql As String, ByVal options() As String) As String
        Dim s As String = ""
        Dim i As Integer
        Dim n As Integer = 0

        If options Is Nothing Or UBound(options, 1) < 0 Then
            s = sql
        Else
            For i = 1 To Len(sql)
                If Mid(sql, i, 1) = "?" Then
                    If n <= UBound(options, 1) Then
                        s &= options(n)
                        n += 1
                    Else

                    End If
                Else
                    s &= Mid(sql, i, 1)
                End If
            Next
        End If
        Return s

    End Function

    Protected Function loadByPrimaryKey(ByVal key As String) As DataSet

        Dim options(1) As String
        options(0) = key
        Try
            Dim stmt As String = prepareStatement(loadStmt, options)
            Dim da As New OleDb.OleDbDataAdapter(stmt, CN.ConnectionString)
            Dim ds As New DataSet(DS_DATA_SET)
            da.Fill(ds, DS_DATA_SET)

            Return ds
        Catch ex As Exception
            MsgBox("Error while retrieving record from " & localTable & ": " & ex.Message)
            Return Nothing
        End Try

    End Function

    Protected Function loadAll() As DataSet

        Try
            Dim da As New OleDb.OleDbDataAdapter(loadAllStmt, CN.ConnectionString)
            Dim ds As New DataSet(DS_DATA_SET)
            da.Fill(ds, DS_DATA_SET)
            Return ds
        Catch ex As Exception
            MsgBox("Error while retrieving records from " & localTable & ": " & ex.Message)
            Return Nothing
        End Try

    End Function

    Protected Function load(ByVal stmt As String, ByVal options() As String) As DataSet

        If Not IsNothing(options) Then
            If Trim(options(0)) <> "" Then
                stmt = prepareStatement(stmt, options)
            End If
        End If
        Dim da As New OleDb.OleDbDataAdapter(stmt, CN.ConnectionString)
        Dim ds As New DataSet(DS_DATA_SET)
        da.Fill(ds, DS_DATA_SET)

        Return ds

    End Function

    Protected Function updateRecord(ByVal key As String, ByVal values() As String) As Boolean
        Try
            Dim cmd As New OleDb.OleDbCommand
            cmd.CommandType = CommandType.Text
            cmd.Connection = CN
            cmd.CommandText = prepareStatement(Me.updateStmt, values)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error while updating record: " & ex.Message)
            Return False
        End Try
        Return True
    End Function

    Protected Function insertRecord(ByVal values() As String) As Integer
        Dim id As Integer = -1
        Try
            Dim da As New OleDb.OleDbDataAdapter(loadAllStmt, CN.ConnectionString)
            Dim ds As New DataSet(DS_DATA_SET)
            da.Fill(ds, DS_DATA_SET)
            'Dim cmd As New OleDb.OleDbCommandBuilder(da)
            'cmd.QuotePrefix = "[" : cmd.QuoteSuffix = "]"
            'da.InsertCommand = cmd.getinsertcommand

            Dim cmd As New OleDb.OleDbCommand
            cmd.CommandType = CommandType.Text
            cmd.Connection = CN
            cmd.CommandText = prepareStatement(Me.insertStmt, values)
            'Dim parm As OleDb.OleDbParameter = New OleDb.OleDbParameter("@newID", OleDb.OleDbType.Integer, 4, priKeyName)
            'parm.Direction = ParameterDirection.Output
            'cmd.Parameters.Add(parm)
            cmd.ExecuteNonQuery()
            'id = CType(parm.Value, Integer)
            id = 0
        Catch ex As OleDb.OleDbException
            If InStr(LCase(ex.Message), "duplicate") > 0 Then
                MsgBox("Attempt to add a duplicate record.")
            Else
                MsgBox("Error while inserting record into " & localTable & ": " & ex.Message)
            End If
            Return id
        Catch ex As Exception
            MsgBox("Error while inserting record into " & localTable & ": " & ex.Message)
            Return id
        End Try
        Return id
    End Function

    Public Function delete(ByVal key As String) As Boolean

        Try
            Dim values() As String = {key}
            Dim cmd As New OleDb.OleDbCommand
            cmd.CommandType = CommandType.Text
            cmd.Connection = CN
            cmd.CommandText = prepareStatement(SQL_DELETE, values)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error while deleting record: " & ex.Message)
            Return False
        End Try
        Return True

    End Function
End Class
