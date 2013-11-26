Imports System.IO
Imports Microsoft.Data.ConnectionUI
Imports Microsoft.Data.ConnectionUI.Sample.VB.Microsoft.Data.ConnectionUI.VB
Imports System.Data.SqlClient
Imports MbcsUtils

Module Main

    Public Const NAME_SPACES As String = "namespaces"
    Public Const GLOBAL_NAMES As String = "globalnames"
    public config As MbcsUtils.DictUtils
    Public configs As Dictionary(Of String, String)
    Public configFile As String = "DbTableClass.config"
    Public nameSpaces As String()
    public globalNames As String()
    Public db As String = Nothing
    Public connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
    Public sqlConnectionStringMbcsCentral As String = "Data Source=MIKE-PC\SQLEXPRESSMBCS;Initial Catalog=MbcsCentral;Integrated Security=True"
    Public sqlConnectionStringMbcsMain As String = "Data Source=MIKE-PC\SQLEXPRESSMBCS;Initial Catalog=MbcsMain;Integrated Security=True"
    Public isSQL As Boolean = False
    Public CN As System.Data.OleDb.OleDbConnection
    Public sqlCN As System.Data.SqlClient.SqlConnection
    Public Const COL_PREFIX As String = "SQL_COL_"
    Public Const PROPERTY_PREFIX As String = "m_"
    Public Const DB_PATH As String = "DatabasePath"
    Public tables() As Table
    Public listTables As ArrayList
    Public chkChanged As Boolean = False
    Public chkProperties As Boolean = False
    Public databasePath As String = Nothing
    Public selectedPath As String = ""
    Dim frmP As Progress

    Public Sub main(ByVal args() As String)

        If args.Length > 0 Then
            selectedPath = args(0).Replace(Chr(34), "")
        End If
        Dim frm As New frmChooseType
        Try
            frm.ShowDialog()
        Catch ex As SystemException
        End Try

    End Sub

    Public Sub process()

        Try
            CN = New System.Data.OleDb.OleDbConnection(db)
            CN.Open()
        Catch ex As System.Data.OleDb.OleDbException
            MsgBox("Cannot open Access Database: " & ex.Message)
            Exit Sub
        End Try
        Dim c As Integer = 0
        frmP = New Progress
        frmP.Show()
        frmP.lst.Items.Add("< < Creating Tables Array > >")
        For Each itm As ListViewItem In listTables
            If itm.Checked Then
                ReDim Preserve tables(c)
                tables(c) = New Table(itm.Text, CN)
                frmP.lst.Items.Add("  " & itm.Text)
                For Each col As columns In tables(c).getCols
                    frmP.lst.Items.Add(String.Format("     {0} / {1} / {2}", col.getName, col.getDbType, col.getIsKey))
                Next
                c += 1
            End If
        Next

        frmP.lst.Items.Add(" ")
        frmP.lst.Items.Add("< < Creating Classes > >")
        For Each Table As Table In tables
            buildClass(Table)
        Next

    End Sub

    Private Sub buildClass(ByVal table As Table)

        frmP.lst.Items.Add(" ")
        frmP.lst.Items.Add(" CLASS " & IIf(table.getName.Contains("_"), table.getName.Substring(3), table.getName))

        Dim key As String = Nothing, keyCol As String = Nothing, keyVar As String = Nothing

        Try
            Using fs As New StreamWriter(String.Format("{0}\{1}.vb", selectedPath, table.getName), False)
                Dim tw As TextWriter = fs
                tw.WriteLine("Imports System.Data")
                If chkChanged Then
                    tw.WriteLine("Imports System.ComponentModel")
                End If
                If isSQL Then
                    tw.WriteLine("Imports System.Data.SqlClient")
                Else
                    tw.WriteLine("Imports System.Data.OleDb" & vbCrLf)
                End If
                tw.WriteLine("Public Class " & table.getName)
                tw.WriteLine(vbTab & "Inherits clsDB")
                If chkChanged Then
                    tw.WriteLine(vbTab & "Implements INotifyPropertyChanged")
                End If
                tw.WriteLine()
                If chkChanged Then
                    tw.WriteLine(vbTab & "Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged")
                    tw.WriteLine()
                End If
                tw.WriteLine(String.Format("{0}Shared TABLE As String = ""{1}""", vbTab, table.getName))
                tw.WriteLine()
                Dim cols() As columns = table.getCols

                ' SQL Column Name constants
                frmP.lst.Items.Add("  Column Constants")
                For Each col As columns In cols
                    tw.WriteLine(String.Format("{0}Public Const {1}{2} AS String = ""{3}""", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, col.getName.Substring(2)))
                    If col.getIsKey Then
                        If Not IsNothing(key) Then
                            key &= " AND "
                        Else
                            keyCol = COL_PREFIX & setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper
                            keyVar = col.getName.Replace(" ", "").Replace("-", "")
                        End If
                        key &= String.Format("{0}{1} & "" = {2}""", COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(col.getDbType.Contains("String"), "'?';", "?;"))
                    End If
                    frmP.lst.Items.Add("        " & String.Format("{0}Public Const {1}{2} AS String = ""{3}""", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, col.getName.Substring(2)))
                Next
                tw.WriteLine()

                ' LOAD Statement
                frmP.lst.Items.Add("  LOAD Statement")
                Dim c As Integer = cols.Length - 1
                Dim n As Integer = 0
                tw.WriteLine(String.Format("{0}Shared SQL_LOAD As String = "" SELECT "" & _", vbTab))
                For Each col As columns In cols
                    tw.WriteLine(String.Format("{0}{0}{0}""["" & {1}{2} & ""]{3}"" & _", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(n < c, ",", "")))
                    n += 1
                Next
                tw.WriteLine(String.Format("{0}{0}{0}"" FROM "" & TABLE & _", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}"" WHERE "" & {1}", vbTab, key))
                tw.WriteLine()

                ' LOAD_ALL Statement
                frmP.lst.Items.Add("  LOAD_ALL Statement")
                n = 0
                tw.WriteLine(String.Format("{0}Shared SQL_LOAD_ALL As String = "" SELECT "" & _", vbTab))
                For Each col As columns In cols
                    tw.WriteLine(String.Format("{0}{0}{0}""["" & {1}{2} & ""]{3}"" & _", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(n < c, ",", "")))
                    n += 1
                Next
                tw.WriteLine(String.Format("{0}{0}{0}"" FROM "" & TABLE & "";""", vbTab))
                tw.WriteLine()

                ' INSERT Statement
                frmP.lst.Items.Add("  INSERT Statement")
                n = 0
                tw.WriteLine(String.Format("{0}Shared SQL_INSERT As String = "" INSERT INTO "" & TABLE & "" ("" & _", vbTab))
                For Each col As columns In cols
                    If Not col.getIsAutoNumber Then
                        tw.WriteLine(String.Format("{0}{0}{0}""["" & {1}{2} & ""]{3}"" & _", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(n < c - 1, ",", "")))
                        n += 1
                    End If
                Next
                tw.Write(String.Format("{0}{0}{0}"") VALUES (", vbTab))
                n = 0
                For Each col As columns In cols
                    If Not col.getIsAutoNumber Then
                        If col.getDbType.Contains("String") Or col.getDbType.Contains("var") Then
                            tw.Write("'?'")
                        ElseIf col.getDbType.Contains("Date") Then
                            tw.Write("#?#")
                        Else
                            tw.Write("?")
                        End If
                        If n < c - 1 Then
                            tw.Write(", ")
                        End If
                        n += 1
                    End If
                Next
                tw.WriteLine(");""")
                tw.WriteLine()

                ' UPDATE Statement
                frmP.lst.Items.Add("  UPDATE Statement")
                n = 0
                tw.WriteLine(String.Format("{0}Shared SQL_UPDATE As String = "" UPDATE "" & TABLE & "" SET "" & _", vbTab))
                For Each col As columns In cols
                    If Not col.getIsAutoNumber Then
                        tw.WriteLine(String.Format("{0}{0}{0}""["" & {1}{2} & ""]={3}{4}"" & _", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(col.getDbType.Contains("String") Or col.getDbType.Contains("var"), "'?'", IIf(col.getDbType.Contains("Date"), "#?#", "?")), IIf(n < c - 1, ",", "")))
                        n += 1
                    End If
                Next
                tw.WriteLine(String.Format("{0}{0}{0}"" WHERE "" & {1}", vbTab, key))
                tw.WriteLine()

                ' Fields
                frmP.lst.Items.Add("  FIELDS")
                For Each col In cols
                    tw.WriteLine(String.Format("{0}Protected {1} As {2}", vbTab, col.getName.Replace(" ", "").Replace("-", ""), (col)))
                Next
                If chkChanged Then
                    tw.WriteLine(vbTab & "Protected m_isChanged As Boolean")
                End If
                tw.WriteLine()

                ' Empty Constructor
                frmP.lst.Items.Add("  Empty New Statement")
                tw.WriteLine(vbTab & "Public Sub New()")
                tw.WriteLine(String.Format("{0}{0}MyBase.New(TABLE, " & keyCol & ", SQL_LOAD, SQL_LOAD_ALL, SQL_INSERT, SQL_UPDATE, CN)", vbTab))
                tw.WriteLine(vbTab & "End Sub")
                tw.WriteLine()

                ' Normal Constructor
                frmP.lst.Items.Add("  Normal New Statement")
                n = 0
                tw.Write(vbTab & "Public Sub New(")
                For Each col As columns In cols
                    tw.Write(String.Format("ByVal {0} As {1}{2}", col.getName.Replace(" ", "").Replace("-", "").Substring(2), col.getDbType.Substring(7), IIf(n < c, ",", "")))
                    n += 1
                Next
                tw.WriteLine(")")
                tw.WriteLine(String.Format("{0}{0}MyBase.New(TABLE, " & keyCol & ", SQL_LOAD, SQL_LOAD_ALL, SQL_INSERT, SQL_UPDATE, CN)", vbTab))
                For Each col In cols
                    tw.WriteLine(String.Format("{0}{0}{1} = {2}", vbTab, col.getName.Replace(" ", "").Replace("-", ""), col.getName.Replace(" ", "").Replace("-", "").Substring(2)))
                Next
                tw.WriteLine(vbTab & "End Sub")
                tw.WriteLine()

                ' fill Statement
                frmP.lst.Items.Add("  fill Statement")
                tw.WriteLine(String.Format("{0}Protected Function fill(ByVal dr as DataRow) As {1}", vbTab, table.getName))
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}Try", vbTab))
                For Each col As columns In cols
                    If col.getDbType.ToLower.Contains("string") Or col.getDbType.ToLower.Contains("text") Then
                        tw.WriteLine(String.Format("{0}{0}{0}{1} = iif(isDBnull(dr({2}{3})),"""",dr({2}{3}))", vbTab, col.getName.Replace(" ", "").Replace("-", ""), COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper))
                    ElseIf col.getDbType.ToLower.Contains("date") Then
                        tw.WriteLine(String.Format("{0}{0}{0}{1} = iif(isDBnull(dr({2}{3})),""#01/01/1900#"",dr({2}{3}))", vbTab, col.getName.Replace(" ", "").Replace("-", ""), COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper))
                    Else
                        tw.WriteLine(String.Format("{0}{0}{0}{1} = dr({2}{3})", vbTab, col.getName.Replace(" ", "").Replace("-", ""), COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper))
                    End If
                Next
                tw.WriteLine(String.Format("{0}{0}Catch ex As {1}", vbTab, IIf(isSQL, "SqlException", "OleDbException")))
                tw.WriteLine(String.Format("{0}{0}Catch exg as System.Exception", vbTab))
                tw.WriteLine(String.Format("{0}{0}End Try", vbTab))
                tw.WriteLine(String.Format("{0}{0}Return Me", vbTab))
                tw.WriteLine()
                tw.WriteLine(vbTab & "End Function")
                tw.WriteLine()

                ' load by id Statement
                frmP.lst.Items.Add("  load by id Statement")
                Dim started As Boolean = False
                tw.Write(vbTab & "Public Overloads Shared Function load(")
                For Each col As columns In cols
                    If col.getIsKey Then
                        tw.Write(String.Format("{0}ByVal {1} As {2}", IIf(started, ", ", ""), col.getName.Replace(" ", "").Replace("-", ""), col.getDbType.Substring(7)))
                        started = True
                    End If
                Next
                tw.WriteLine(") As " & table.getName)
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}Dim obj As {1} = New {1}", vbTab, table.getName))
                tw.Write(String.Format("{0}{0}Dim opt() as String = {{", vbTab))
                started = False
                For Each col As columns In cols
                    If col.getIsKey Then
                        tw.Write(String.Format("{0}{1}", IIf(started, ", ", ""), col.getName.Replace(" ", "").Replace("-", "")))
                        started = True
                    End If
                Next
                tw.WriteLine("}")
                tw.WriteLine(String.Format("{0}{0}Dim ds As DataSet = New DataSet", vbTab))
                tw.WriteLine(String.Format("{0}{0}ds = obj.load(SQL_LOAD, opt)", vbTab))
                tw.WriteLine(String.Format("{0}{0}Dim dr as DataRow = ds.Tables(0).Rows(0)", vbTab))
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}Return obj.fill(dr)", vbTab))
                tw.WriteLine()
                tw.WriteLine(vbTab & "End Function")
                tw.WriteLine()

                ' load Statement
                frmP.lst.Items.Add("  load Statement")
                tw.WriteLine(String.Format("{0}Public Overloads Shared Function load() As {1}()", vbTab, table.getName))
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}Dim list as ArrayList = New ArrayList()", vbTab))
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}Dim ds As DataSet = New DataSet", vbTab))
                tw.WriteLine(String.Format("{0}{0}Dim obj As {1} = New {1}", vbTab, table.getName))
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}ds = obj.loadAll()", vbTab))
                tw.WriteLine(String.Format("{0}{0}obj = Nothing", vbTab))
                tw.WriteLine(String.Format("{0}{0}Dim dr as DataRow", vbTab))
                tw.WriteLine(String.Format("{0}{0}For Each dr in ds.Tables(0).Rows", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}obj = New {1}", vbTab, table.getName))
                tw.WriteLine(String.Format("{0}{0}{0}obj.fill(dr)", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}list.Add(obj)", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}obj = Nothing", vbTab))
                tw.WriteLine(String.Format("{0}{0}Next", vbTab))
                tw.WriteLine(String.Format("{0}{0}Try", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}Return DirectCast(list.ToArray(GetType({1})), {1}())", vbTab, table.getName))
                tw.WriteLine(String.Format("{0}{0}Catch ex As {1}", vbTab, IIf(isSQL, "SqlException", "OleDbException")))
                tw.WriteLine(String.Format("{0}{0}{0}Return Nothing", vbTab))
                tw.WriteLine(String.Format("{0}{0}End Try", vbTab))
                tw.WriteLine()
                tw.WriteLine(vbTab & "End Function")
                tw.WriteLine()

                ' save Statement
                frmP.lst.Items.Add("  save Statment")
                tw.WriteLine(vbTab & "Public Function save() As Boolean")
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}If " & keyVar & " > 0 Then", vbTab))
                tw.Write(String.Format("{0}{0}{0}Dim values() as String = {{", vbTab))
                started = False
                For Each col As columns In cols
                    If Not col.getIsAutoNumber Then
                        tw.Write(CType(IIf(started, ", ", ""), String) & col.getName.Replace(" ", "").Replace("-", ""))
                        started = True
                    End If
                Next
                For Each col As columns In cols
                    If col.getIsKey Then
                        tw.Write(", " & col.getName.Replace(" ", "").Replace("-", ""))
                    End If
                Next
                tw.WriteLine("}")
                tw.WriteLine(String.Format("{0}{0}{0}If Me.updateRecord(" & keyVar & ", values) Then", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}{0}Return True", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}Else", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}{0}Return False", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}End If", vbTab))
                tw.WriteLine(String.Format("{0}{0}Else", vbTab))
                tw.Write(String.Format("{0}{0}{0}Dim values() as String = {{", vbTab))
                started = False
                For Each col As columns In cols
                    If Not col.getIsAutoNumber Then
                        tw.Write(CType(IIf(started, ", ", ""), String) & col.getName.Replace(" ", "").Replace("-", ""))
                        started = True
                    End If
                Next
                tw.WriteLine("}")
                tw.WriteLine(String.Format("{0}{0}{0}" & keyVar & " = Me.insertRecord(values)", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}If " & keyVar & " >= 0 Then", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}{0}Return True", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}End If", vbTab))
                tw.WriteLine(String.Format("{0}{0}End If", vbTab))
                tw.WriteLine(String.Format("{0}{0}Return False", vbTab))
                tw.WriteLine()
                tw.WriteLine(vbTab & "End Function")
                tw.WriteLine()

                ' Create Table method
                frmP.lst.Items.Add("  CreateTable Method")
                frmP.lst.Items.Add("      " & buildCreateTable(table, cols))
                tw.WriteLine()
                tw.WriteLine(vbTab & "Public Shared Function CreateTable(CN as System.Data.OleDb.OleDbConnection) as Boolean")
                tw.WriteLine()
                tw.WriteLine(String.Format("{0}{0}Try", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}CN.Open()", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}Dim cmd As New OleDb.OleDbCommand", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}cmd.CommandType = CommandType.Text", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}cmd.CommandText=""{1}""", vbTab, buildCreateTable(table, cols)))
                tw.WriteLine(String.Format("{0}{0}{0}cmd.Connection = CN", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}cmd.ExecuteNonQuery()", vbTab))
                tw.WriteLine(String.Format("{0}{0}Catch ex as {1}", vbTab, IIf(isSQL, "SqlException", "OleDbException")))
                tw.WriteLine(String.Format("{0}{0}{0}Return False", vbTab))
                tw.WriteLine(String.Format("{0}{0}Finally", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}Try", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}{0}CN.Close()", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}Catch", vbTab))
                tw.WriteLine(String.Format("{0}{0}{0}End Try", vbTab))
                tw.WriteLine(String.Format("{0}{0}End Try", vbTab))
                tw.WriteLine()
                tw.WriteLine("{0}{0}Return True", vbTab)
                tw.WriteLine()
                tw.WriteLine(vbTab & "End Function")
                tw.WriteLine()

                If chkChanged Then
                    tw.WriteLine(vbTab & "Private Sub OnPropertyChanged(ByVal info As String)")
                    tw.WriteLine(vbTab & vbTab & "RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(info))")
                    tw.WriteLine(String.Format("{0}{0}m_isChanged = True", vbTab))
                    tw.WriteLine(vbTab & "End Sub")
                    tw.WriteLine()
                End If
                If Not chkProperties Then
                    ' Get/Set Statements
                    frmP.lst.Items.Add("  Get/Set Statements")
                    tw.WriteLine(vbTab & "' Get/Set Routines")
                    For Each col As columns In cols
                        frmP.lst.Items.Add("    " & col.getName)
                        If Not col.getIsAutoNumber Then
                            tw.WriteLine(String.Format("{0}Public Sub set{1}(ByVal value As {2})", vbTab, col.getName.Replace(" ", "").Replace("-", "").Substring(2), col.getDbType.Substring(7)))
                            tw.WriteLine(String.Format("{0}{0}{1} = value", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                            tw.WriteLine(vbTab & "End Sub")
                        End If
                        tw.WriteLine(String.Format("{0}Public Function get{1}() As {2}", vbTab, col.getName.Replace(" ", "").Replace("-", "").Substring(2), col.getDbType.Substring(7)))
                        tw.WriteLine(String.Format("{0}{0}Return {1}", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                        tw.WriteLine(vbTab & "End Function")
                        tw.WriteLine()
                    Next
                Else
                    ' Property Statements
                    frmP.lst.Items.Add("  Property Statements")
                    tw.WriteLine(vbTab & "' Property Routines")
                    If chkChanged Then
                        tw.WriteLine(vbTab & "Public Property isChanged As Boolean")
                        tw.WriteLine(vbTab & vbTab & "Get")
                        tw.WriteLine(String.Format("{0}{0}{0}Return m_isChanged", vbTab))
                        tw.WriteLine(vbTab & vbTab & "End Get")
                        tw.WriteLine(vbTab & vbTab & "Set(value As Boolean)")
                        tw.WriteLine(vbTab & vbTab & vbTab & "m_isChanged = value")
                        tw.WriteLine(vbTab & vbTab & "End Set")
                        tw.WriteLine(vbTab & "End Property")
                    End If
                    For Each col As columns In cols
                        frmP.lst.Items.Add("    " & col.getName)
                        tw.WriteLine(String.Format("{0}Public {1} Property {2} As {3}", vbTab, IIf(col.getIsAutoNumber, "ReadOnly", ""), col.getName.Replace(" ", "").Replace("-", "").Substring(2), convertDBType(col)))
                        tw.WriteLine(String.Format("{0}{0}Get", vbTab))
                        tw.WriteLine(String.Format("{0}{0}{0}Return {1}", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                        tw.WriteLine(vbTab & vbTab & "End Get")
                        If Not col.getIsAutoNumber Then
                            tw.WriteLine(String.Format("{0}{0}Set(ByVal value As {1})", vbTab, convertDBType(col)))
                            tw.WriteLine(String.Format("{0}{0}{0}{1} = value", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                            If chkChanged Then tw.WriteLine(String.Format("{0}{0}{1}""{2}""{3}", vbTab, "OnPropertyChanged(", col.getName.Replace(" ", "").Replace("-", "").Substring(2), ")"))
                            tw.WriteLine(String.Format("{0}{0}End Set", vbTab))
                        End If
                        tw.WriteLine(vbTab & "End Property")
                        tw.WriteLine()
                    Next
                End If

                tw.WriteLine("End Class")
                tw.Flush()
                tw.Close()
                fs.Close()
            End Using

        Catch ex As IOException
            MsgBox("Error creating output file: " & ex.Message)
        End Try

    End Sub

    Public Function setConstant(s As String) As String

        Dim x As String = ""
        Dim underScore As Boolean = False, first As Boolean = True
        If s.StartsWith("_") Then
            s = s.Substring(1)
        End If
        For Each c As Char In s
            If c = "_" Then
                underScore = True
            Else
                If Not first And (Char.IsUpper(c) Or underScore) Then
                    x &= "_"
                End If
                x &= c
                underScore = False
            End If
            first = False
        Next
        Return x

    End Function

    Function buildCreateTable(table As Table, cols() As columns) As String

        Dim stmt As String = "CREATE TABLE " & table.getName & " ("
        For Each col As columns In cols
            stmt &= col.getName.Substring(2) & " " & getDataType(col) & ","
        Next
        stmt = stmt.Substring(0, stmt.Length - 1) & ")"
        Return stmt

    End Function

    Function getDataType(col As columns) As String

        Dim t As String = ""
        Dim c As String = col.getDbType.ToLower
        If c.Contains("int32") Or c.Contains("integer") Or c.Contains("bigint") Then
            If col.getIsAutoNumber Then
                t = "AUTONUMBER"
            Else
                t = "LONG"
            End If
        ElseIf c.Contains("int16") Or c.Contains("int") Then
            t = "INTEGER"
        ElseIf c.Contains("boolean") Then
            t = "YESNO"
        ElseIf c.Contains("string") Or c.Contains("varchar") Or c.Contains("sysname") Then
            If col.getColumnSize <= 255 Then
                t = "TEXT(" & col.getColumnSize.ToString & ")"
            Else
                t = "MEMO"
            End If
        ElseIf c.Contains("date") Then
            t = "DATETIME"
        ElseIf c.Contains("single") Then
            t = "SINGLE"
        ElseIf c.Contains("double") Then
            t = "DOUBLE"
        ElseIf c.Contains("currency") Or c.Contains("decimal") Or c.Contains("float") Then
            t = "DECIMAL"
        End If
        Return t

    End Function

    Public Function convertDBType(col As columns) As String

        Dim t As String = ""
        Dim c As String = col.getDbType.ToLower
        If c.Contains("int32") Or c.Contains("integer") Or c.Contains("bigint") Then
            t = "Long"
        ElseIf c.Contains("int") Then
            t = "Integer"
        ElseIf c.Contains("boolean") Or c.Contains("bit") Then
            t = "Boolean"
        ElseIf c.Contains("string") Or c.Contains("varchar") Or c.Contains("sysname") Then
            t = "String"
        ElseIf c.Contains("date") Then
            t = "Date"
        ElseIf c.Contains("single") Or c.Contains("double") Then
            t = c
        ElseIf c.Contains("currency") Or c.Contains("decimal") Or c.Contains("float") Then
            t = "Decimal"
        ElseIf c.Contains("binary") Then
            t = "Object"
        End If
        Return t

    End Function

    Public Sub SetDbConnection(db As Integer)


        'Dim dcd As New DataConnectionDialog()
        'Dim dcs As New DataConnectionConfiguration(Nothing)
        'dcs.LoadConfiguration(dcd)

        'dcd.ConnectionString = SqlConnectionString

        'If DataConnectionDialog.Show(dcd) = DialogResult.OK Then
        ' load tables 
        'Using connection As New SqlConnection(dcd.ConnectionString)
        'connection.Open()
        'Dim cmd As New SqlCommand("SELECT * FROM sys.Tables", connection)

        'Using reader As SqlDataReader = cmd.ExecuteReader()
        'While reader.Read()
        'Console.WriteLine(reader.HasRows)
        'End While
        'sqlConnectionString = dcd.ConnectionString
        'End Using
        'connection.Close()
        'End Using
        If IsNothing(sqlCN) Then
            sqlCN = New System.Data.SqlClient.SqlConnection
        End If
        Select Case db
            Case 1 : sqlCN.ConnectionString = sqlConnectionStringMbcsCentral
            Case 2 : sqlCN.ConnectionString = sqlConnectionStringMbcsMain
        End Select
        'End If
        'dcs.SaveConfiguration(dcd)

    End Sub

End Module
