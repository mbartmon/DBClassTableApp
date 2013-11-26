Imports System.IO

Public Class frmMain
    Const COL_PREFIX As String = "SQL_COL_"
    ' Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
    Dim db As String = Nothing
    Public CN As System.Data.OleDb.OleDbConnection
    Public targetPath As String = Nothing
    Dim tables() As Table
    Dim frm As Progress
    Private Sub cmdSource_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSource.Click

        OpenFileDialog1.Filter = "mdb|*.mdb;*.accdb"
        OpenFileDialog1.InitialDirectory = ""
        OpenFileDialog1.ShowDialog()

        Dim fileName As String = Nothing

        fileName = OpenFileDialog1.FileName

        lstTables.Clear()
        If Not IsNothing(fileName) Then
            txtSource.Text = fileName
            db = connectionString & fileName
            fillTableList()
        End If
    End Sub

    Sub fillTableList()

        Dim s As String = Nothing
        Try
            CN = New System.Data.OleDb.OleDbConnection(db)
            CN.Open()
        Catch ex As System.Data.OleDb.OleDbException
            MsgBox("Cannot open Access Database: " & ex.Message)
            Exit Sub
        End Try

        Dim lstItem As ListViewItem
        Dim dt As DataTable = CN.GetSchema("Tables")
        For Each dr As DataRow In dt.Rows
            If dr.Item(3).Equals("TABLE") Then
                s = dr.Item(2)
                lstItem = lstTables.Items.Add(s)
                'lstItem.Checked = True
            End If
        Next
        CN.Close()

    End Sub

    Private Sub cmdDestination_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDestination.Click

        FolderBrowserDialog1.SelectedPath = "c:\workfiles"
        FolderBrowserDialog1.ShowDialog()

        targetPath = FolderBrowserDialog1.SelectedPath
        txtDestination.Text = targetPath

    End Sub

    Private Sub cmdProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdProcess.Click

        Try
            CN = New System.Data.OleDb.OleDbConnection(db)
            CN.Open()
        Catch ex As System.Data.OleDb.OleDbException
            MsgBox("Cannot open Access Database: " & ex.Message)
            Exit Sub
        End Try
        Dim c As Integer = 0
        frm = New Progress
        frm.Show()
        frm.lst.Items.Add("< < Creating Tables Array > >")
        For Each itm As ListViewItem In lstTables.Items
            If itm.Checked Then
                ReDim Preserve tables(c)
                tables(c) = New Table(itm.Text, CN)
                frm.lst.Items.Add("  " & itm.Text)
                For Each col As columns In tables(c).getCols
                    frm.lst.Items.Add(String.Format("     {0} / {1} / {2}", col.getName, col.getDbType, col.getIsKey))
                Next
                c += 1
            End If
        Next

        frm.lst.Items.Add(" ")
        frm.lst.Items.Add("< < Creating Classes > >")
        For Each Table As Table In tables
            buildClass(Table)
        Next

    End Sub

    Private Sub buildClass(ByVal table As Table)

        frm.lst.Items.Add(" ")
        frm.lst.Items.Add(" CLASS " & table.getName)

        Dim key As String = Nothing, keyCol As String = Nothing, keyVar As String = Nothing
        Dim fsAlt As StreamWriter

        Using fs As New StreamWriter(String.Format("{0}\{1}.vb", targetPath, table.getName), False)
            Dim tw As TextWriter = fs
            tw.WriteLine("Imports System.Data")
            If chkChanged.Checked Then
                tw.WriteLine("Imports System.ComponentModel")
            End If
            tw.WriteLine("Imports System.Data.OleDb" & vbCrLf)
            tw.WriteLine("Public Class " & table.getName)
            tw.WriteLine(vbTab & "Inherits clsDB")
            If chkChanged.Checked Then
                tw.WriteLine(vbTab & "Implements INotifyPropertyChanged")
            End If
            tw.WriteLine()
            If chkChanged.Checked Then
                tw.WriteLine(vbTab & "Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged")
                tw.WriteLine()
            End If
            tw.WriteLine(String.Format("{0}Shared TABLE As String = ""{1}""", vbTab, table.getName))
            tw.WriteLine()
            Dim cols() As columns = table.getCols
            ' SQL Column Name constants
            frm.lst.Items.Add("  Column Constants")
            For Each col As columns In cols
                tw.WriteLine(String.Format("{0}Public Const {1}{2} AS String = ""{3}""", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, col.getName.Substring(2)))
                If col.getIsKey Then
                    If Not IsNothing(key) Then
                        key &= " AND "
                    Else
                        keyCol = COL_PREFIX & setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper
                        keyVar = col.getName.Replace(" ", "").Replace("-", "")
                    End If
                    key &= String.Format("""["" & {0}{1} & ""] = {2}""", COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(col.getDbType.Contains("String"), "'?';", "?;"))
                End If
                frm.lst.Items.Add("        " & String.Format("{0}Public Const {1}{2} AS String = ""{3}""", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, col.getName.Substring(2)))
            Next
            tw.WriteLine()

            ' LOAD Statement
            frm.lst.Items.Add("  LOAD Statement")
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
            frm.lst.Items.Add("  LOAD_ALL Statement")
            n = 0
            tw.WriteLine(String.Format("{0}Shared SQL_LOAD_ALL As String = "" SELECT "" & _", vbTab))
            For Each col As columns In cols
                tw.WriteLine(String.Format("{0}{0}{0}""["" & {1}{2} & ""]{3}"" & _", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(n < c, ",", "")))
                n += 1
            Next
            tw.WriteLine(String.Format("{0}{0}{0}"" FROM "" & TABLE & "";""", vbTab))
            tw.WriteLine()

            ' INSERT Statement
            frm.lst.Items.Add("  INSERT Statement")
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
                    If col.getDbType.Contains("String") Then
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
            frm.lst.Items.Add("  UPDATE Statement")
            n = 0
            tw.WriteLine(String.Format("{0}Shared SQL_UPDATE As String = "" UPDATE "" & TABLE & "" SET "" & _", vbTab))
            For Each col As columns In cols
                If Not col.getIsAutoNumber Then
                    tw.WriteLine(String.Format("{0}{0}{0}""["" & {1}{2} & ""]={3}{4}"" & _", vbTab, COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(col.getDbType.Contains("String"), "'?'", IIf(col.getDbType.Contains("Date"), "#?#", "?")), IIf(n < c - 1, ",", "")))
                    n += 1
                End If
            Next
            tw.WriteLine(String.Format("{0}{0}{0}"" WHERE "" & {1}", vbTab, key))
            tw.WriteLine()

            ' Fields
            frm.lst.Items.Add("  FIELDS")
            For Each col In cols
                tw.WriteLine(String.Format("{0}Protected {1} As {2}", vbTab, col.getName.Replace(" ", "").Replace("-", ""), col.getDbType.Substring(7)))
            Next
            tw.WriteLine()
            ' Empty Constructor
            frm.lst.Items.Add("  Empty New Statement")
            tw.WriteLine(vbTab & "Public Sub New()")
            tw.WriteLine(String.Format("{0}{0}MyBase.New(TABLE, " & keyCol & ", SQL_LOAD, SQL_LOAD_ALL, SQL_INSERT, SQL_UPDATE, CN)", vbTab))
            tw.WriteLine(vbTab & "End Sub")
            tw.WriteLine()
            ' Normal Constructor
            frm.lst.Items.Add("  Normal New Statement")
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
            frm.lst.Items.Add("  fill Statement")
            tw.WriteLine(String.Format("{0}Protected Function fill(ByVal dr as DataRow) As {1}", vbTab, table.getName))
            tw.WriteLine()
            tw.WriteLine(String.Format("{0}{0}Try", vbTab))
            For Each col As columns In cols
                tw.WriteLine(String.Format("{0}{0}{0}{1} = dr({2}{3})", vbTab, col.getName.Replace(" ", "").Replace("-", ""), COL_PREFIX, setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper))
            Next
            tw.WriteLine(String.Format("{0}{0}Catch ex As OleDbException", vbTab))
            tw.WriteLine(String.Format("{0}{0}End Try", vbTab))
            tw.WriteLine(String.Format("{0}{0}Return Me", vbTab))
            tw.WriteLine()
            tw.WriteLine(vbTab & "End Function")
            tw.WriteLine()
            ' load by id Statement
            frm.lst.Items.Add("  load by id Statement")
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
            frm.lst.Items.Add("  load Statement")
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
            tw.WriteLine(String.Format("{0}{0}Catch ex As OleDbException", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}Return Nothing", vbTab))
            tw.WriteLine(String.Format("{0}{0}End Try", vbTab))
            tw.WriteLine()
            tw.WriteLine(vbTab & "End Function")
            tw.WriteLine()
            ' save Statement
            frm.lst.Items.Add("  save Statment")
            tw.WriteLine(vbTab & "Public Function save() As Boolean")
            tw.WriteLine()
            tw.WriteLine(String.Format("{0}{0}If " & keyVar & " > 0 Then", vbTab))
            tw.Write(String.Format("{0}{0}{0}Dim values() as String = {{", vbTab))
            started = False
            For Each col As columns In cols
                If Not col.getIsAutoNumber Then
                    tw.Write(IIf(started, ", ", "") & col.getName.Replace(" ", "").Replace("-", ""))
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
                    tw.Write(IIf(started, ", ", "") & col.getName.Replace(" ", "").Replace("-", ""))
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

            ' ALTER TABLE Statements
            If chkAlter.Checked Then
                fsAlt = New StreamWriter(String.Format("{0}\{1}.cs", targetPath, table.getName & "_alter"), False)
                Dim twa As TextWriter = fsAlt

                frm.lst.Items.Add("  Alter Table Method")
                twa.WriteLine(vbTab & "public static void alterTable()")
                twa.WriteLine(vbTab & "{")
                twa.WriteLine(vbTab & vbTab & "System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();")
                twa.WriteLine(vbTab & vbTab & "cmd.CommandType = CommandType.Text;")
                twa.WriteLine(vbTab & vbTab & "cmd.Connection = GlobalShared.CN;")
                For Each col As columns In cols
                    twa.WriteLine(vbTab & vbTab & "try {")
                    twa.WriteLine(String.Format("{0}{0}{0}cmd.CommandText = ""ALTER TABLE "" & TABLE & "" ADD [{1}] {2}"";", vbTab, col.getName.Replace("-", "").Substring(2), getSQLDataType(col)))
                    twa.WriteLine(vbTab & vbTab & vbTab & "cmd.ExecuteNonQuery();")
                    twa.WriteLine(vbTab & vbTab & "} catch (System.Data.SqlClient.SqlException ex) {")
                    twa.WriteLine(vbTab & vbTab & "} catch {System.Exception ex) {")
                    twa.WriteLine(vbTab & vbTab & "}")
                    twa.WriteLine(vbTab & vbTab & "try {")
                    twa.WriteLine(vbTab & vbTab & "GlobalShared.CN.Close();")
                    twa.WriteLine(vbTab & vbTab & "} catch {}")
                Next
                twa.WriteLine(vbTab & "}")
                twa.WriteLine()
                twa.Flush()
                fsAlt.Flush()
                twa.Close()
                fsAlt.Close()
                twa = Nothing
                fsAlt = Nothing
            End If

            If chkChanged.Checked Then
                tw.WriteLine(vbTab & "Private Sub OnPropertyChanged(ByVal info As String)")
                tw.WriteLine(vbTab & vbTab & "RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(info))")
                tw.WriteLine(vbTab & "End Sub")
                tw.WriteLine()
            End If

            If Not chkProperties.Checked Then
                ' Get/Set Statements
                frm.lst.Items.Add("  Get/Set Statements")
                tw.WriteLine(vbTab & "' Get/Set Routines")
                For Each col As columns In cols
                    frm.lst.Items.Add("    " & col.getName)
                    If Not col.getIsAutoNumber Then
                        tw.WriteLine(String.Format("{0}Public Sub set{1}(ByVal value As {2})", vbTab, col.getName.Replace(" ", "").Replace("-", ""), col.getDbType.Substring(7)))
                        tw.WriteLine(String.Format("{0}{0}{1} = value", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                        tw.WriteLine(vbTab & "End Sub")
                    End If
                    tw.WriteLine(String.Format("{0}Public Function get{1}() As {2}", vbTab, col.getName.Replace(" ", "").Replace("-", ""), col.getDbType.Substring(7)))
                    tw.WriteLine(String.Format("{0}{0}Return {1}", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                    tw.WriteLine(vbTab & "End Function")
                    tw.WriteLine()
                Next
            Else
                ' Property Statements
                frm.lst.Items.Add("  Property Statements")
                tw.WriteLine(vbTab & "' Property Routines")
                For Each col As columns In cols
                    frm.lst.Items.Add("    " & col.getName)
                    tw.WriteLine(String.Format("{0}Public {1} Property {2} As {3}", vbTab, IIf(col.getIsAutoNumber, "ReadOnly", ""), col.getName.Replace(" ", "").Replace("-", "").Substring(2), col.getDbType.Substring(7)))
                    tw.WriteLine(String.Format("{0}{0}Get", vbTab))
                    tw.WriteLine(String.Format("{0}{0}{0}Return {1}", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                    tw.WriteLine(vbTab & vbTab & "End Get")
                    If Not col.getIsAutoNumber Then
                        tw.WriteLine(String.Format("{0}{0}Set(ByVal value As {1})", vbTab, col.getDbType.Substring(7)))
                        tw.WriteLine(String.Format("{0}{0}{0}{1} = value", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
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

    End Sub


    Shared Function getSQLDataType(col As columns) As String

        Dim t As String = ""
        Dim c As String = col.getDbType.ToLower
        If c.Contains("int32") Or c.Contains("integer") Or c.Contains("bigint") Then
            t = "bigint"
        ElseIf c.Contains("int16") Or c.Contains("int") Then
            t = "int"
        ElseIf c.Contains("boolean") Then
            t = "bit"
        ElseIf c.Contains("string") Or c.Contains("varchar") Or c.Contains("sysname") Then
            If col.getColumnSize <= 255 Then
                t = "varchar(" & col.getColumnSize.ToString & ")"
            Else
                t = "varchar(MAX)"
            End If
        ElseIf c.Contains("date") Then
            t = "datetime2(0)"
        ElseIf c.Contains("single") Then
            t = "real"
        ElseIf c.Contains("double") Then
            t = "real"
        ElseIf c.Contains("currency") Or c.Contains("decimal") Or c.Contains("float") Then
            t = "money"
        End If
        Return t

    End Function

    Function setConstant(s As String) As String

        Dim x As String = ""
        For Each c As Char In s
            If Char.IsUpper(c) Then
                x &= "_"
            End If
            x &= c
        Next
        Return x

    End Function

    Private Sub chkProperties_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkProperties.CheckedChanged

        If Not chkProperties.Checked Then
            chkChanged.Checked = False
        End If

    End Sub

    Private Sub cmdSelectAll_Click(sender As System.Object, e As System.EventArgs) Handles cmdSelectAll.Click

        For Each itm As ListViewItem In lstTables.Items
            itm.Checked = True
        Next

    End Sub

    Private Sub cmdUnselect_Click(sender As Object, e As System.EventArgs) Handles cmdUnselect.Click

        For Each itm As ListViewItem In lstTables.Items
            itm.Checked = False
        Next

    End Sub

    Private Sub cmdExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub
End Class
