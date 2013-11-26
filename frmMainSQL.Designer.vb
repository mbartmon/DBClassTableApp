<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMainSQL
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtSource = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdSource = New System.Windows.Forms.Button()
        Me.lstTables = New System.Windows.Forms.ListView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmdProcess = New System.Windows.Forms.Button()
        Me.cmdDestination = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDestination = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.chkProperties = New System.Windows.Forms.CheckBox()
        Me.chkChanged = New System.Windows.Forms.CheckBox()
        Me.cmdSelectAll = New System.Windows.Forms.Button()
        Me.cmdUnselect = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optC = New System.Windows.Forms.RadioButton()
        Me.optVB = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtNameSpace = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtGlobal = New System.Windows.Forms.TextBox()
        Me.lstNameSpaces = New System.Windows.Forms.ComboBox()
        Me.lstGlobalNames = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "mdb"
        Me.OpenFileDialog1.Filter = "mdb|*.mdb"
        Me.OpenFileDialog1.InitialDirectory = "c:\pmsys"
        '
        'txtSource
        '
        Me.txtSource.Location = New System.Drawing.Point(90, 12)
        Me.txtSource.Name = "txtSource"
        Me.txtSource.Size = New System.Drawing.Size(227, 20)
        Me.txtSource.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Access DB"
        '
        'cmdSource
        '
        Me.cmdSource.Location = New System.Drawing.Point(339, 10)
        Me.cmdSource.Name = "cmdSource"
        Me.cmdSource.Size = New System.Drawing.Size(75, 23)
        Me.cmdSource.TabIndex = 2
        Me.cmdSource.Text = "BROWSE"
        Me.cmdSource.UseVisualStyleBackColor = True
        '
        'lstTables
        '
        Me.lstTables.CheckBoxes = True
        Me.lstTables.Location = New System.Drawing.Point(12, 92)
        Me.lstTables.Name = "lstTables"
        Me.lstTables.Size = New System.Drawing.Size(402, 265)
        Me.lstTables.TabIndex = 3
        Me.lstTables.UseCompatibleStateImageBehavior = False
        Me.lstTables.View = System.Windows.Forms.View.List
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Tables"
        '
        'cmdProcess
        '
        Me.cmdProcess.Location = New System.Drawing.Point(288, 421)
        Me.cmdProcess.Name = "cmdProcess"
        Me.cmdProcess.Size = New System.Drawing.Size(75, 23)
        Me.cmdProcess.TabIndex = 5
        Me.cmdProcess.Text = "PROCESS"
        Me.cmdProcess.UseVisualStyleBackColor = True
        '
        'cmdDestination
        '
        Me.cmdDestination.Location = New System.Drawing.Point(339, 46)
        Me.cmdDestination.Name = "cmdDestination"
        Me.cmdDestination.Size = New System.Drawing.Size(75, 23)
        Me.cmdDestination.TabIndex = 8
        Me.cmdDestination.Text = "BROWSE"
        Me.cmdDestination.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 51)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Destination"
        '
        'txtDestination
        '
        Me.txtDestination.Location = New System.Drawing.Point(90, 48)
        Me.txtDestination.Name = "txtDestination"
        Me.txtDestination.Size = New System.Drawing.Size(227, 20)
        Me.txtDestination.TabIndex = 6
        '
        'chkProperties
        '
        Me.chkProperties.AutoSize = True
        Me.chkProperties.Checked = True
        Me.chkProperties.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkProperties.Location = New System.Drawing.Point(15, 367)
        Me.chkProperties.Name = "chkProperties"
        Me.chkProperties.Size = New System.Drawing.Size(101, 17)
        Me.chkProperties.TabIndex = 9
        Me.chkProperties.Text = "Use Properties?"
        Me.chkProperties.UseVisualStyleBackColor = True
        '
        'chkChanged
        '
        Me.chkChanged.AutoSize = True
        Me.chkChanged.Checked = True
        Me.chkChanged.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkChanged.Location = New System.Drawing.Point(122, 367)
        Me.chkChanged.Name = "chkChanged"
        Me.chkChanged.Size = New System.Drawing.Size(136, 17)
        Me.chkChanged.TabIndex = 10
        Me.chkChanged.Text = "Use PropertyChanged?"
        Me.chkChanged.UseVisualStyleBackColor = True
        '
        'cmdSelectAll
        '
        Me.cmdSelectAll.Location = New System.Drawing.Point(67, 72)
        Me.cmdSelectAll.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdSelectAll.Name = "cmdSelectAll"
        Me.cmdSelectAll.Size = New System.Drawing.Size(56, 19)
        Me.cmdSelectAll.TabIndex = 11
        Me.cmdSelectAll.Text = "Select All"
        Me.cmdSelectAll.UseVisualStyleBackColor = True
        '
        'cmdUnselect
        '
        Me.cmdUnselect.Location = New System.Drawing.Point(128, 72)
        Me.cmdUnselect.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdUnselect.Name = "cmdUnselect"
        Me.cmdUnselect.Size = New System.Drawing.Size(77, 19)
        Me.cmdUnselect.TabIndex = 12
        Me.cmdUnselect.Text = "Unselect All"
        Me.cmdUnselect.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(362, 421)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(52, 23)
        Me.cmdExit.TabIndex = 13
        Me.cmdExit.Text = "EXIT"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optC)
        Me.GroupBox1.Controls.Add(Me.optVB)
        Me.GroupBox1.Location = New System.Drawing.Point(265, 367)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(149, 46)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Output Language"
        '
        'optC
        '
        Me.optC.AutoSize = True
        Me.optC.Checked = True
        Me.optC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optC.Location = New System.Drawing.Point(74, 19)
        Me.optC.Name = "optC"
        Me.optC.Size = New System.Drawing.Size(41, 17)
        Me.optC.TabIndex = 17
        Me.optC.TabStop = True
        Me.optC.Text = "C#"
        Me.optC.UseVisualStyleBackColor = True
        '
        'optVB
        '
        Me.optVB.AutoSize = True
        Me.optVB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optVB.Location = New System.Drawing.Point(6, 19)
        Me.optVB.Name = "optVB"
        Me.optVB.Size = New System.Drawing.Size(41, 17)
        Me.optVB.TabIndex = 16
        Me.optVB.Text = "VB"
        Me.optVB.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 391)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 13)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Namespace"
        '
        'txtNameSpace
        '
        Me.txtNameSpace.Location = New System.Drawing.Point(90, 388)
        Me.txtNameSpace.Name = "txtNameSpace"
        Me.txtNameSpace.Size = New System.Drawing.Size(168, 20)
        Me.txtNameSpace.TabIndex = 16
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 441)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(37, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Global"
        '
        'txtGlobal
        '
        Me.txtGlobal.Location = New System.Drawing.Point(90, 438)
        Me.txtGlobal.Name = "txtGlobal"
        Me.txtGlobal.Size = New System.Drawing.Size(168, 20)
        Me.txtGlobal.TabIndex = 18
        '
        'lstNameSpaces
        '
        Me.lstNameSpaces.FormattingEnabled = True
        Me.lstNameSpaces.Location = New System.Drawing.Point(91, 409)
        Me.lstNameSpaces.Name = "lstNameSpaces"
        Me.lstNameSpaces.Size = New System.Drawing.Size(168, 21)
        Me.lstNameSpaces.TabIndex = 19
        '
        'lstGlobalNames
        '
        Me.lstGlobalNames.FormattingEnabled = True
        Me.lstGlobalNames.Location = New System.Drawing.Point(90, 457)
        Me.lstGlobalNames.Name = "lstGlobalNames"
        Me.lstGlobalNames.Size = New System.Drawing.Size(168, 21)
        Me.lstGlobalNames.TabIndex = 20
        '
        'frmMainSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(427, 490)
        Me.Controls.Add(Me.lstGlobalNames)
        Me.Controls.Add(Me.lstNameSpaces)
        Me.Controls.Add(Me.txtGlobal)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtNameSpace)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdUnselect)
        Me.Controls.Add(Me.cmdSelectAll)
        Me.Controls.Add(Me.chkChanged)
        Me.Controls.Add(Me.chkProperties)
        Me.Controls.Add(Me.cmdDestination)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDestination)
        Me.Controls.Add(Me.cmdProcess)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lstTables)
        Me.Controls.Add(Me.cmdSource)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSource)
        Me.Name = "frmMainSQL"
        Me.Text = "Create DB Table Classes"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtSource As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdSource As System.Windows.Forms.Button
    Friend WithEvents lstTables As System.Windows.Forms.ListView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdProcess As System.Windows.Forms.Button
    Friend WithEvents cmdDestination As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDestination As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents chkProperties As System.Windows.Forms.CheckBox
    Friend WithEvents chkChanged As System.Windows.Forms.CheckBox
    Friend WithEvents cmdSelectAll As System.Windows.Forms.Button
    Friend WithEvents cmdUnselect As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optC As System.Windows.Forms.RadioButton
    Friend WithEvents optVB As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtNameSpace As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtGlobal As System.Windows.Forms.TextBox
    Friend WithEvents lstNameSpaces As System.Windows.Forms.ComboBox
    Friend WithEvents lstGlobalNames As System.Windows.Forms.ComboBox

End Class
