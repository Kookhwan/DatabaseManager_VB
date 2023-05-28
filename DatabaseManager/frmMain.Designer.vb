<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.dgvDatabase = New System.Windows.Forms.DataGridView()
        Me.pbProgress = New System.Windows.Forms.ProgressBar()
        Me.tabDatabase = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnCreateBackup1 = New System.Windows.Forms.Button()
        Me.cmbDatabase1 = New System.Windows.Forms.ComboBox()
        Me.lblDatabase1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnCreateBackup2 = New System.Windows.Forms.Button()
        Me.cmbDatabase2 = New System.Windows.Forms.ComboBox()
        Me.lblDatabase2 = New System.Windows.Forms.Label()
        Me.dgvProgress = New System.Windows.Forms.DataGridView()
        Me.chkWebSQL = New System.Windows.Forms.CheckBox()
        CType(Me.dgvDatabase, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabDatabase.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDatabase
        '
        Me.dgvDatabase.AllowUserToAddRows = False
        Me.dgvDatabase.AllowUserToDeleteRows = False
        Me.dgvDatabase.AllowUserToOrderColumns = True
        Me.dgvDatabase.AllowUserToResizeRows = False
        Me.dgvDatabase.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDatabase.Location = New System.Drawing.Point(12, 200)
        Me.dgvDatabase.MultiSelect = False
        Me.dgvDatabase.Name = "dgvDatabase"
        Me.dgvDatabase.ReadOnly = True
        Me.dgvDatabase.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDatabase.Size = New System.Drawing.Size(437, 207)
        Me.dgvDatabase.TabIndex = 4
        '
        'pbProgress
        '
        Me.pbProgress.Location = New System.Drawing.Point(12, 413)
        Me.pbProgress.Name = "pbProgress"
        Me.pbProgress.Size = New System.Drawing.Size(776, 25)
        Me.pbProgress.TabIndex = 5
        '
        'tabDatabase
        '
        Me.tabDatabase.Controls.Add(Me.TabPage1)
        Me.tabDatabase.Controls.Add(Me.TabPage2)
        Me.tabDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabDatabase.Location = New System.Drawing.Point(12, 12)
        Me.tabDatabase.Name = "tabDatabase"
        Me.tabDatabase.SelectedIndex = 0
        Me.tabDatabase.Size = New System.Drawing.Size(441, 182)
        Me.tabDatabase.TabIndex = 8
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.AliceBlue
        Me.TabPage1.Controls.Add(Me.btnCreateBackup1)
        Me.TabPage1.Controls.Add(Me.cmbDatabase1)
        Me.TabPage1.Controls.Add(Me.lblDatabase1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(433, 153)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "By Live Database"
        '
        'btnCreateBackup1
        '
        Me.btnCreateBackup1.Location = New System.Drawing.Point(236, 84)
        Me.btnCreateBackup1.Name = "btnCreateBackup1"
        Me.btnCreateBackup1.Size = New System.Drawing.Size(175, 44)
        Me.btnCreateBackup1.TabIndex = 17
        Me.btnCreateBackup1.Text = "Create a backup by DB"
        Me.btnCreateBackup1.UseVisualStyleBackColor = True
        '
        'cmbDatabase1
        '
        Me.cmbDatabase1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDatabase1.FormattingEnabled = True
        Me.cmbDatabase1.Location = New System.Drawing.Point(114, 20)
        Me.cmbDatabase1.Name = "cmbDatabase1"
        Me.cmbDatabase1.Size = New System.Drawing.Size(297, 23)
        Me.cmbDatabase1.TabIndex = 16
        '
        'lblDatabase1
        '
        Me.lblDatabase1.AutoSize = True
        Me.lblDatabase1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDatabase1.Location = New System.Drawing.Point(6, 23)
        Me.lblDatabase1.Name = "lblDatabase1"
        Me.lblDatabase1.Size = New System.Drawing.Size(102, 15)
        Me.lblDatabase1.TabIndex = 15
        Me.lblDatabase1.Text = "Source Database"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Lavender
        Me.TabPage2.Controls.Add(Me.btnCreateBackup2)
        Me.TabPage2.Controls.Add(Me.cmbDatabase2)
        Me.TabPage2.Controls.Add(Me.lblDatabase2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(433, 153)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "By backed up file"
        '
        'btnCreateBackup2
        '
        Me.btnCreateBackup2.Location = New System.Drawing.Point(236, 84)
        Me.btnCreateBackup2.Name = "btnCreateBackup2"
        Me.btnCreateBackup2.Size = New System.Drawing.Size(175, 44)
        Me.btnCreateBackup2.TabIndex = 18
        Me.btnCreateBackup2.Text = "Create a backup by file"
        Me.btnCreateBackup2.UseVisualStyleBackColor = True
        '
        'cmbDatabase2
        '
        Me.cmbDatabase2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDatabase2.FormattingEnabled = True
        Me.cmbDatabase2.Location = New System.Drawing.Point(114, 20)
        Me.cmbDatabase2.Name = "cmbDatabase2"
        Me.cmbDatabase2.Size = New System.Drawing.Size(297, 23)
        Me.cmbDatabase2.TabIndex = 17
        '
        'lblDatabase2
        '
        Me.lblDatabase2.AutoSize = True
        Me.lblDatabase2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDatabase2.Location = New System.Drawing.Point(37, 23)
        Me.lblDatabase2.Name = "lblDatabase2"
        Me.lblDatabase2.Size = New System.Drawing.Size(71, 15)
        Me.lblDatabase2.TabIndex = 16
        Me.lblDatabase2.Text = "Source files"
        '
        'dgvProgress
        '
        Me.dgvProgress.AllowUserToAddRows = False
        Me.dgvProgress.AllowUserToDeleteRows = False
        Me.dgvProgress.AllowUserToOrderColumns = True
        Me.dgvProgress.AllowUserToResizeRows = False
        Me.dgvProgress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProgress.Location = New System.Drawing.Point(455, 37)
        Me.dgvProgress.MultiSelect = False
        Me.dgvProgress.Name = "dgvProgress"
        Me.dgvProgress.ReadOnly = True
        Me.dgvProgress.RowHeadersVisible = False
        Me.dgvProgress.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvProgress.Size = New System.Drawing.Size(333, 370)
        Me.dgvProgress.TabIndex = 9
        '
        'chkWebSQL
        '
        Me.chkWebSQL.AutoSize = True
        Me.chkWebSQL.Checked = True
        Me.chkWebSQL.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkWebSQL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkWebSQL.Location = New System.Drawing.Point(622, 12)
        Me.chkWebSQL.Name = "chkWebSQL"
        Me.chkWebSQL.Size = New System.Drawing.Size(166, 19)
        Me.chkWebSQL.TabIndex = 20
        Me.chkWebSQL.Text = "Create a Web Test DB"
        Me.chkWebSQL.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.chkWebSQL)
        Me.Controls.Add(Me.dgvProgress)
        Me.Controls.Add(Me.tabDatabase)
        Me.Controls.Add(Me.pbProgress)
        Me.Controls.Add(Me.dgvDatabase)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Database Manager"
        CType(Me.dgvDatabase, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabDatabase.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.dgvProgress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgvDatabase As DataGridView
    Friend WithEvents pbProgress As ProgressBar
    Friend WithEvents tabDatabase As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents cmbDatabase1 As ComboBox
    Friend WithEvents lblDatabase1 As Label
    Friend WithEvents btnCreateBackup1 As Button
    Friend WithEvents cmbDatabase2 As ComboBox
    Friend WithEvents lblDatabase2 As Label
    Friend WithEvents btnCreateBackup2 As Button
    Friend WithEvents dgvProgress As DataGridView
    Friend WithEvents chkWebSQL As CheckBox
End Class
