Public Class frmUserInterface

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub cmdSource_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSource.Click
        Dim obj As Object = sender
        OpenFileDialog1.Filter = "mdb|*.mdb;*.accdb"
        OpenFileDialog1.InitialDirectory = databasePath
        OpenFileDialog1.ShowDialog()

        databasePath = OpenFileDialog1.FileName
        lstTables.Clear()
        If Not IsNothing(databasePath) Then
            txtSource.Text = databasePath
            db = connectionString & databasePath
            fillTableList()
            prj.Globals(DB_PATH) = databasePath
            prj.Globals.VariablePersists(DB_PATH) = True
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
                s = dr.Item(2).ToString
                lstItem = lstTables.Items.Add(s)
                'lstItem.Checked = True
            End If
        Next
        CN.Close()

    End Sub

    Private Sub cmdDestination_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDestination.Click

        FolderBrowserDialog1.SelectedPath = selectedPath
        FolderBrowserDialog1.ShowDialog()

        selectedPath = FolderBrowserDialog1.SelectedPath
        txtDestination.Text = selectedPath

    End Sub

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

    Private Sub cmdExit_Click(sender As System.Object, e As System.EventArgs)
        Me.Visible = False
        Me.Hide()
    End Sub

    Private Sub cmdProcess_Click(sender As System.Object, e As System.EventArgs) Handles cmdProcess.Click

        listTables = New ArrayList
        For Each itm As ListViewItem In lstTables.Items
            listTables.Add(itm)
        Next
        Main.chkChanged = Me.chkChanged.Checked
        Main.chkProperties = Me.chkProperties.Checked
        Main.process()
    End Sub

    Private Sub frmUserInterface_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.Width = 445
        Me.Height = 500
        txtDestination.Text = selectedPath
        txtSource.Text = databasePath
    End Sub
End Class
