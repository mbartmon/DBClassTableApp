Imports System.IO
Imports System.Data.SqlClient

Public Class frmMainSQL
    Const COL_PREFIX As String = "SQL_COL_"
    ' Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
    Dim db As String = Nothing
    Public CN As System.Data.OleDb.OleDbConnection
    Public targetPath As String = Nothing
    Dim tables() As Table
    Dim frm As Progress

    Private Sub frmMainSQL_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        fillTableList()

        targetPath = IIf(Main.selectedPath.Trim() = "", "c:\workfiles", Main.selectedPath)
        txtDestination.Text = targetPath

        configs = New Dictionary(Of String, String)
        config = New MbcsUtils.DictUtils(Application.StartupPath & "\" & configFile, configs)
        config.getConfigs()
        Dim s As String = ""
        config.getItem(NAME_SPACES, s)
        If Not IsNothing(s) AndAlso s.Length > 0 Then
            nameSpaces = s.Split(",")
        Else
            ReDim nameSpaces(0)
        End If
        config.getItem(GLOBAL_NAMES, s)
        If Not IsNothing(s) AndAlso s.Length > 0 Then
            globalNames = s.Split(",")
        Else
            ReDim globalNames(0)
        End If
        For Each s In nameSpaces
            lstNameSpaces.Items.Add(s)
        Next
        For Each s In globalNames
            lstGlobalNames.Items.Add(s)
        Next

    End Sub

    Private Sub frmMainSQL_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        dim s As String = ""
        If lstNameSpaces.Items.Count > 0 Then
            For i As Integer = 0 To lstNameSpaces.Items.Count - 1
                s &= lstNameSpaces.Items(i)
                If i < lstNameSpaces.Items.Count - 1 Then
                    s &= ","
                End If
            Next
            config.setItem(NAME_SPACES, s)
            s = ""
        End If
        If lstGlobalNames.Items.Count > 0 Then
            For i As Integer = 0 To lstGlobalNames.Items.Count - 1
                s &= lstGlobalNames.Items(i)
                If i < lstGlobalNames.Items.Count - 1 Then
                    s &= ","
                End If
            Next
            config.setItem(GLOBAL_NAMES, s)
        End If
        config.setConfigs()
        config = Nothing

    End Sub

    Private Sub cmdSource_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSource.Click

        OpenFileDialog1.Filter = "mdb|*.mdb;*.accdb"
        OpenFileDialog1.InitialDirectory = ""
        OpenFileDialog1.ShowDialog()

        Dim fileName As String = Nothing

        fileName = OpenFileDialog1.FileName

        lstTables.Clear()
        If Not IsNothing(fileName) Then
            fillTableList()
        End If
    End Sub

    Sub fillTableList()

        Dim s As String = Nothing
        Try
            sqlCN.Open()
        Catch ex As SqlException
            MsgBox("Cannot open SQL SERVER Database: " & ex.Message)
            Exit Sub
        End Try

        Dim lstItem As ListViewItem
        Dim da As New SqlClient.SqlDataAdapter("SELECT * FROM sys.Tables ORDER BY name", sqlCN)
        Dim ds As New DataSet("Tables")
        da.Fill(ds, "Tables")
        For Each dr As DataRow In ds.Tables(0).Rows
            s = dr("name")
            lstItem = lstTables.Items.Add(s)
        Next
        sqlCN.Close()

    End Sub

    Private Sub cmdDestination_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDestination.Click

        FolderBrowserDialog1.SelectedPath = "c:\workfiles"
        FolderBrowserDialog1.ShowDialog()

        targetPath = FolderBrowserDialog1.SelectedPath
        txtDestination.Text = targetPath

    End Sub

    Private Sub cmdProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdProcess.Click

        If optVB.Checked Then
            SqlBuildVB.Process(lstTables, targetPath, chkChanged.Checked, chkProperties.Checked)
        ElseIf optC.Checked Then
            SqlBuildCSharp.Process(lstTables, targetPath, chkChanged.Checked, chkProperties.Checked, txtNameSpace.Text, txtGlobal.Text)
        Else
            MsgBox("No output language selected.")
        End If
        Return
    End Sub


    'Public Shared Function setConstant(s As String) As String

    'Dim x As String = ""
    'Dim lastWasUnderscore As Boolean = False
    '   For Each c As Char In s
    '       If c.Equals("_") then lastWasUnderscore = True
    '      If Char.IsUpper(c) And Not lastWasUnderscore Then
    '         x &= "_"
    '    End If
    '   x &= c
    '  Next
    '  Return x

    ' End Function

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

    Private Sub txtNameSpace_LostFocus(sender As Object, e As System.EventArgs) Handles txtNameSpace.LostFocus

        Dim found As Boolean = False
        For Each s As String In lstNameSpaces.Items
            If txtNameSpace.Text.ToLower.Equals(s.ToLower) Then
                found = True
            End If
        Next
        If Not found Then
            lstNameSpaces.Items.Add(txtNameSpace.Text)
        End If

    End Sub

    Private Sub txtGlobal_LostFocus(sender As Object, e As System.EventArgs) Handles txtGlobal.LostFocus

        Dim found As Boolean = False
        For Each s As String In lstGlobalNames.Items
            If txtGlobal.Text.ToLower.Equals(s.ToLower) Then
                found = True
            End If
        Next
        If Not found Then
            lstGlobalNames.Items.Add(txtGlobal.Text)
        End If

    End Sub

    Private Sub lstNameSpaces_SelectedValueChanged1(sender As Object, e As System.EventArgs) Handles lstNameSpaces.SelectedValueChanged

        txtNameSpace.Text = lstNameSpaces.SelectedItem

    End Sub

    Private Sub lstGlobalNames_SelectedValueChanged1(sender As Object, e As System.EventArgs) Handles lstGlobalNames.SelectedValueChanged

        txtGlobal.Text = lstGlobalNames.SelectedItem

    End Sub
End Class
