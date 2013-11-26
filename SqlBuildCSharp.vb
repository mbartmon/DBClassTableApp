Imports System.IO
Imports System.Data.SqlClient

Public Class SqlBuildCSharp

    Shared frm As Progress
    Shared namespaceName As String = ""
    Shared namspaceFlag As Boolean=False
    Shared globalName As String = ""
    Shared globalFlag As Boolean = False

    Public Shared Sub Process(lstTables As ListView, targetPath As String, changed As Boolean, properties As Boolean, namespaceN As String, globalN As String)

        Try
            sqlCN.Open()
        Catch ex As SqlException
            MsgBox("Cannot open SQL SERVER Database: " & ex.Message)
            Exit Sub
        End Try
        If Not namespaceN.Trim().Equals("") Then
            namspaceFlag=True
            namespaceName = namespaceN.Trim()
        End If
        If Not globalN.Trim().Equals("") Then
            globalFlag = True
            globalName = globalN.Trim()
        End If
        Dim c As Integer = 0
        frm = New Progress
        frm.Show()
        frm.lst.Items.Add("< < Creating Tables Array > >")
        For Each itm As ListViewItem In lstTables.Items
            If itm.Checked Then
                ReDim Preserve tables(c)
                tables(c) = New Table(itm.Text, sqlCN)
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
            buildClass(Table, targetPath, changed, properties)
        Next
        sqlCN.Close()

    End Sub

    Shared Sub buildClass(ByVal table As Table, targetPath As String, changed As Boolean, properties As Boolean)

        frm.lst.Items.Add(" ")
        frm.lst.Items.Add(" CLASS " & IIf(table.getName.Contains("_"), table.getName.Substring(3), table.getName))

        Dim key As String = Nothing, keyCol As String = Nothing, keyVar As String = Nothing
        Dim className As String = IIf(table.getName.Contains("_"), table.getName.Substring(3), table.getName)

        Using fs As New StreamWriter(String.Format("{0}\{1}.cs", targetPath, className, False))
            Dim tw As TextWriter = fs
            tw.WriteLine("using System;")
            tw.WriteLine("using System.Data;")
            tw.WriteLine("using System.Collections;")
            If changed Then
                tw.WriteLine("using System.ComponentModel;")
            End If
            tw.WriteLine("using System.Data.SqlClient;")
            tw.WriteLine()
            If namspaceFlag Then
                tw.WriteLine("namespace " & namespaceName & " {")
            End If
            tw.Write("public class " & className & " : clsDB")
            If changed Then
                tw.Write(vbTab & ", INotifyPropertyChanged")
            End If
            tw.WriteLine() : tw.WriteLine("{")
            If changed Then
                tw.WriteLine(vbTab & "public event PropertyChangedEventHandler PropertyChanged;")
                tw.WriteLine()
            End If
            tw.WriteLine(String.Format("{0}public static string TABLE = ""{1}"";", vbTab, table.getName))
            tw.WriteLine()
            Dim cols() As columns = table.getCols

            ' SQL Column Name constants
            frm.lst.Items.Add("  Column Constants")
            For Each col As columns In cols
                tw.WriteLine(String.Format("{0}public const string {1}{2} = ""{3}"";", vbTab, COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, col.getName.Substring(2)))
                If col.getIsKey Then
                    If Not IsNothing(key) Then
                        key &= " AND "
                    Else
                        keyCol = COL_PREFIX & Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper
                        keyVar = col.getName.Replace(" ", "").Replace("-", "")
                    End If
                    key &= String.Format("[{0}{1}] + "" = {2}""", COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(col.getDbType.Contains("String") Or col.getDbType.Contains("char"), "'?';", "?;"))
                End If
                frm.lst.Items.Add("        " & String.Format("{0}public const string {1}{2} = ""{3}"";", vbTab, COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, col.getName.Substring(2)))
            Next
            tw.WriteLine()

            ' LOAD Statement
            frm.lst.Items.Add("  LOAD Statement")
            Dim c As Integer = cols.Length - 1
            Dim n As Integer = 0
            tw.WriteLine(String.Format("{0}static string SQL_LOAD = ""SELECT "" +", vbTab))
            For Each col As columns In cols
                tw.WriteLine(String.Format("{0}{0}{0}""["" + {1}{2} + ""]{3}"" +", vbTab, COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(n < c, ",", "")))
                n += 1
            Next
            tw.WriteLine(String.Format("{0}{0}{0}"" FROM "" + TABLE +", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}"" WHERE "" + {1} = ?;", vbTab, key))
            tw.WriteLine()

            ' LOAD_ALL Statement
            frm.lst.Items.Add("  LOAD_ALL Statement")
            n = 0
            tw.WriteLine(String.Format("{0}static string SQL_LOAD_ALL = ""SELECT "" +", vbTab))
            For Each col As columns In cols
                tw.WriteLine(String.Format("{0}{0}{0}""["" + {1}{2} + ""]{3}"" +", vbTab, COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(n < c, ",", "")))
                n += 1
            Next
            tw.WriteLine(String.Format("{0}{0}{0}"" FROM "" + TABLE + "";"";", vbTab))
            tw.WriteLine()

            ' INSERT Statement
            frm.lst.Items.Add("  INSERT Statement")
            n = 0
            tw.WriteLine(String.Format("{0}static string SQL_INSERT = ""INSERT INTO "" + TABLE + "" ("" +", vbTab))
            For Each col As columns In cols
                If Not col.getIsAutoNumber Then
                    tw.WriteLine(String.Format("{0}{0}{0}""["" + {1}{2} + ""]{3}"" +", vbTab, COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(n < c - 1, ",", "")))
                    n += 1
                End If
            Next
            tw.Write(String.Format("{0}{0}{0}"") VALUES (", vbTab))
            n = 0
            For Each col As columns In cols
                If Not col.getIsAutoNumber Then
                    If convertDBType(col).Contains("string") Or col.getDbType.Contains("char") Then
                        tw.Write("'?'")
                    ElseIf col.getDbType.ToLower().Contains("date") Then
                        tw.Write("'?'")
                    Else
                        tw.Write("?")
                    End If
                    If n < c - 1 Then
                        tw.Write(", ")
                    End If
                    n += 1
                End If
            Next
            tw.WriteLine(");"";")
            tw.WriteLine()

            ' UPDATE Statement
            frm.lst.Items.Add("  UPDATE Statement")
            n = 0
            tw.WriteLine(String.Format("{0}static string SQL_UPDATE = ""UPDATE "" + TABLE + "" SET "" +", vbTab))
            For Each col As columns In cols
                If Not col.getIsAutoNumber Then
                    tw.WriteLine(String.Format("{0}{0}{0}""["" + {1}{2} + ""]={3}{4}"" +", vbTab, COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper, IIf(convertDBType(col).Contains("String") Or col.getDbType.Contains("char"), "'?'", IIf(convertDBType(col).Contains("Date"), "'?'", "?")), IIf(n < c - 1, ",", "")))
                    n += 1
                End If
            Next
            tw.WriteLine(String.Format("{0}{0}{0}"" WHERE "" + {1};", vbTab, key))
            tw.WriteLine()

            ' Fields
            frm.lst.Items.Add("  FIELDS")
            For Each col In cols
                tw.WriteLine(String.Format("{0}protected {2} {1};", vbTab, col.getName.Replace(" ", "").Replace("-", ""), convertDBType(col)))
            Next
            If changed Then
                tw.WriteLine(vbTab & "protected bool m_isChanged;")
            End If
            tw.WriteLine()

            ' Empty Constructor
            frm.lst.Items.Add("  Empty New Statement")
            tw.Write(vbTab & "public " & className & "() : ")
            tw.WriteLine("base(TABLE, " & keyCol & ", SQL_LOAD, SQL_LOAD_ALL, SQL_INSERT, SQL_UPDATE, " & IIf(globalFlag, globalName & ".", "") & "CN)")
            tw.WriteLine(vbTab & "{")
            tw.WriteLine(vbTab & "}")
            tw.WriteLine()

            ' Normal Constructor
            frm.lst.Items.Add("  Normal New Statement")
            n = 0
            tw.Write(vbTab & "public " & className & "(")
            For Each col As columns In cols
                tw.Write(String.Format("{1} {0}{2}", col.getName.Replace(" ", "").Replace("-", "").Substring(2), convertDBType(col), IIf(n < cols.Length - 1, ",", "")))
                n += 1
            Next
            tw.WriteLine(") : ")
            tw.WriteLine(vbTab & vbTab & vbTab & "base(TABLE, " & keyCol & ", SQL_LOAD, SQL_LOAD_ALL, SQL_INSERT, SQL_UPDATE, " & IIf(globalFlag, globalName & ".", "") & "CN)")
            tw.WriteLine(vbTab & "{")
            For Each col In cols
                tw.WriteLine(String.Format("{0}{0}{1} = {2};", vbTab, col.getName.Replace(" ", "").Replace("-", ""), col.getName.Replace(" ", "").Replace("-", "").Substring(2)))
            Next
            tw.WriteLine(vbTab & "}")
            tw.WriteLine()

            ' fill Statement
            frm.lst.Items.Add("  fill Statement")
            tw.WriteLine(String.Format("{0}protected " & className & " fill(DataRow dr)", vbTab))
            tw.WriteLine(vbTab & "{")
            tw.WriteLine(String.Format("{0}{0}try {{", vbTab))
            For Each col As columns In cols
                If convertDBType(col).ToLower.Contains("string") Or convertDBType(col).ToLower.Contains("text") Then
                    tw.WriteLine(String.Format("{0}{0}{0}{1} = (string) (Convert.IsDBNull(dr[{2}{3}]) ? """" : dr[{2}{3}]);", vbTab, col.getName.Replace(" ", "").Replace("-", ""), COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper))
                ElseIf convertDBType(col).ToLower.Contains("date") Then
                    tw.WriteLine(String.Format("{0}{0}{0}{1} = (DateTime) (Convert.IsDBNull(dr[{2}{3}]) ? new DateTime(1900,1,1,0,0,0) : dr[{2}{3}]);", vbTab, col.getName.Replace(" ", "").Replace("-", ""), COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper))
                Else
                    tw.WriteLine(String.Format("{0}{0}{0}{2} = ({1}) dr[{3}{4}];", vbTab, convertDBType(col), col.getName.Replace(" ", "").Replace("-", ""), COL_PREFIX, Main.setConstant(col.getName.Replace(" ", "_").Replace("-", "_").Substring(2)).ToUpper))
                End If
            Next
            tw.WriteLine(String.Format("{0}{0}}} catch ({1} ex) {{", vbTab, "SqlException"))
            tw.WriteLine(String.Format("{0}{0}}}", vbTab))
            tw.WriteLine(String.Format("{0}{0}return this;", vbTab))
            tw.WriteLine()
            tw.WriteLine(vbTab & "}")
            tw.WriteLine()

            ' load by id Statement
            frm.lst.Items.Add("  load by id Statement")
            Dim started As Boolean = False
            tw.Write(vbTab & "public static " & className & " load(")
            For Each col As columns In cols
                If col.getIsKey Then
                    tw.Write(String.Format("{0} {2} {1}", IIf(started, ", ", ""), col.getName.Replace(" ", "").Replace("-", ""), convertDBType(col)))
                    started = True
                End If
            Next
            tw.WriteLine(")")
            tw.WriteLine(vbTab & "{")
            tw.WriteLine(String.Format("{0}{0}{1} obj  = new {1}();", vbTab, className))
            tw.Write(String.Format("{0}{0}string[] opt = {{", vbTab))
            started = False
            For Each col As columns In cols
                If col.getIsKey Then
                    tw.Write(String.Format("{0}{1}", IIf(started, ", ", ""), col.getName.Replace(" ", "").Replace("-", "") & IIf(Not convertDBType(col).Contains("string"), ".ToString()", "")))
                    started = True
                End If
            Next
            tw.WriteLine("};")
            tw.WriteLine(String.Format("{0}{0}DataSet ds = new DataSet();", vbTab))
            tw.WriteLine(String.Format("{0}{0}ds = obj.load(SQL_LOAD, opt);", vbTab))
            tw.WriteLine(vbTab & vbTab & "if (ds != null) {")
            tw.WriteLine(String.Format("{0}{0}{0}DataRow dr = ds.Tables[0].Rows[0];", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}return obj.fill(dr);", vbTab))
            tw.WriteLine(vbTab & vbTab & "} else")
            tw.WriteLine(vbTab & vbTab & vbTab & "return null;")
            tw.WriteLine()
            tw.WriteLine(vbTab & "}")
            tw.WriteLine()

            ' load All Statement
            frm.lst.Items.Add("  load Statement")
            tw.WriteLine(String.Format("{0}public static {1}[] load()", vbTab, className))
            tw.WriteLine(vbTab & "{")
            tw.WriteLine(String.Format("{0}{0}ArrayList list = new ArrayList();", vbTab))
            tw.WriteLine()
            tw.WriteLine(String.Format("{0}{0}DataSet ds = new DataSet();", vbTab))
            tw.WriteLine(String.Format("{0}{0}{1} obj = new {1}();", vbTab, className))
            tw.WriteLine()
            tw.WriteLine(String.Format("{0}{0}ds = obj.loadAll();", vbTab))
            tw.WriteLine(vbTab & vbTab & "if (ds == null)")
            tw.WriteLine(vbTab & vbTab & vbTab & "return null;")
            tw.WriteLine(String.Format("{0}{0}obj = null;", vbTab))
            tw.WriteLine(String.Format("{0}{0}foreach (DataRow dr in ds.Tables[0].Rows) {{", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}obj = new {1}();", vbTab, className))
            tw.WriteLine(String.Format("{0}{0}{0}obj.fill(dr);", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}list.Add(obj);", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}obj = null;", vbTab))
            tw.WriteLine(String.Format("{0}{0}}}", vbTab))
            tw.WriteLine(String.Format("{0}{0}try {{", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}return ({1}[])list.ToArray(typeof({1}));", vbTab, className))
            tw.WriteLine(String.Format("{0}{0}}} catch ({1} ex) {{", vbTab, "SqlException"))
            tw.WriteLine(String.Format("{0}{0}{0}return null;", vbTab))
            tw.WriteLine(String.Format("{0}{0}}}", vbTab))
            tw.WriteLine()
            tw.WriteLine(vbTab & "}")
            tw.WriteLine()

            ' save Method
            frm.lst.Items.Add("  save Method")
            tw.WriteLine(vbTab & "public bool save()")
            tw.WriteLine(vbTab & "{")
            tw.WriteLine(String.Format("{0}{0}if (" & keyVar & " > 0 ) {{", vbTab))
            tw.Write(String.Format("{0}{0}{0}string[] values = {{", vbTab))
            started = False
            For Each col As columns In cols
                If Not col.getIsAutoNumber Then
                    tw.Write(IIf(started, ", ", "") & col.getName.Replace(" ", "").Replace("-", "") & IIf(Not convertDBType(col).Contains("string"), ".ToString()", ""))
                    started = True
                End If
            Next
            For Each col As columns In cols
                If col.getIsKey Then
                    tw.Write(", " & col.getName.Replace(" ", "").Replace("-", "") & IIf(Not convertDBType(col).Contains("string"), ".ToString()", ""))
                End If
            Next
            tw.WriteLine("};")
            tw.WriteLine(String.Format("{0}{0}{0}if (this.updateRecord(" & keyVar & ".ToString(), values)) {{", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}{0}return true;", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}}} else {{", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}{0}return false;", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}}}", vbTab))
            tw.WriteLine(String.Format("{0}{0}}} else {{", vbTab))
            tw.Write(String.Format("{0}{0}{0}string[] values = {{", vbTab))
            started = False
            For Each col As columns In cols
                If Not col.getIsAutoNumber Then
                    tw.Write(IIf(started, ", ", "") & col.getName.Replace(" ", "").Replace("-", "") & IIf(Not convertDBType(col).Contains("string"), ".ToString()", ""))
                    started = True
                End If
            Next
            tw.WriteLine("};")
            tw.WriteLine(String.Format("{0}{0}{0}" & keyVar & " = this.insertRecord(values);", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}if (" & keyVar & " >= 0) {{", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}{0}return true;", vbTab))
            tw.WriteLine(String.Format("{0}{0}{0}}}", vbTab))
            tw.WriteLine(String.Format("{0}{0}}}", vbTab))
            tw.WriteLine(String.Format("{0}{0}return false;", vbTab))
            tw.WriteLine()
            tw.WriteLine(vbTab & "}")
            tw.WriteLine()

            If changed Then
                tw.WriteLine(vbTab & "private void OnPropertyChanged(string info) {")
                tw.WriteLine(vbTab & vbTab & "if (PropertyChanged != null) {")
                tw.WriteLine(vbTab & vbTab & vbTab & "PropertyChanged(this, new PropertyChangedEventArgs(info));")
                tw.WriteLine(String.Format("{0}{0}{0}m_isChanged = true;", vbTab))
                tw.WriteLine(vbTab & vbTab & "}")
                tw.WriteLine(vbTab & "}")
                tw.WriteLine()
            End If

            ' Property Statements
            frm.lst.Items.Add("  Get/Set Statements")
            tw.WriteLine(vbTab & "// Properties")
            For Each col As columns In cols
                frm.lst.Items.Add("    " & col.getName)
                tw.WriteLine(String.Format("{0}public {1} {2} {{", vbTab, convertDBType(col), col.getName.Replace(" ", "").Replace("-", "").Substring(2)))
                If Not col.getIsAutoNumber Then
                    tw.WriteLine(String.Format("{0}{0}set{{ {1} = value;", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                    If changed Then tw.WriteLine(String.Format("{0}{0}{1}""{2}""{3}", vbTab, "OnPropertyChanged(", col.getName.Replace(" ", "").Replace("-", "").Substring(2), "); }"))
                End If
                tw.WriteLine(String.Format("{0}{0}get {{ return {1}; }}", vbTab, col.getName.Replace(" ", "").Replace("-", "")))
                tw.WriteLine(vbTab & "}")
                tw.WriteLine()
            Next

            tw.WriteLine("}")
            If namspaceFlag Then
                tw.WriteLine("}")
            End If
            tw.Flush()
            tw.Close()
            fs.Close()
        End Using

    End Sub

    Shared Function convertDBType(col As columns) As String

        Dim t As String = ""
        Dim c As String = col.getDbType.ToLower
        If c.Contains("int32") Or c.Contains("integer") Or c.Contains("bigint") Then
            t = "long"
        ElseIf c.Contains("smallint") Then
            t = "short"
        ElseIf c.Contains("int") Then
            t = "int"
        ElseIf c.Contains("real") Then
            t = "float"
        ElseIf c.Contains("float") Then
            t = "float"
        ElseIf c.Contains("boolean") Or c.Contains("bit") Then
            t = "bool"
        ElseIf c.Contains("string") Or c.Contains("varchar") Or c.Contains("sysname") Then
            t = "string"
        ElseIf c.Contains("date") Then
            t = "DateTime"
        ElseIf c.Contains("single") Or c.Contains("double") Then
            t = c
        ElseIf c.Contains("currency") Or c.Contains("decimal") Or c.Contains("float") Or c.Contains("money") Then
            t = "decimal"
        ElseIf c.Contains("binary") Then
            t = "object"
        End If
        Return t

    End Function

End Class
