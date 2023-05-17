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
        Me.dgvDatabase = New System.Windows.Forms.DataGridView()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.tabDatabase = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.cmbDatabase = New System.Windows.Forms.ComboBox()
        Me.lblDatabase = New System.Windows.Forms.Label()
        Me.btnCreateBackup = New System.Windows.Forms.Button()
        CType(Me.dgvDatabase, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabDatabase.SuspendLayout()
        Me.TabPage1.SuspendLayout()
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
        Me.dgvDatabase.Size = New System.Drawing.Size(441, 207)
        Me.dgvDatabase.TabIndex = 4
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 413)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(776, 25)
        Me.ProgressBar1.TabIndex = 5
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(459, 34)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(329, 373)
        Me.DataGridView1.TabIndex = 6
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
        Me.TabPage1.Controls.Add(Me.btnCreateBackup)
        Me.TabPage1.Controls.Add(Me.cmbDatabase)
        Me.TabPage1.Controls.Add(Me.lblDatabase)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(433, 153)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "By new backup file"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Lavender
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(433, 156)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "By existed backup file"
        '
        'cmbDatabase
        '
        Me.cmbDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDatabase.FormattingEnabled = True
        Me.cmbDatabase.Location = New System.Drawing.Point(114, 20)
        Me.cmbDatabase.Name = "cmbDatabase"
        Me.cmbDatabase.Size = New System.Drawing.Size(297, 23)
        Me.cmbDatabase.TabIndex = 16
        '
        'lblDatabase
        '
        Me.lblDatabase.AutoSize = True
        Me.lblDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDatabase.Location = New System.Drawing.Point(6, 23)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(102, 15)
        Me.lblDatabase.TabIndex = 15
        Me.lblDatabase.Text = "Source Database"
        '
        'btnCreateBackup
        '
        Me.btnCreateBackup.Location = New System.Drawing.Point(236, 76)
        Me.btnCreateBackup.Name = "btnCreateBackup"
        Me.btnCreateBackup.Size = New System.Drawing.Size(175, 44)
        Me.btnCreateBackup.TabIndex = 17
        Me.btnCreateBackup.Text = "Create a new backup DB"
        Me.btnCreateBackup.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.tabDatabase)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.dgvDatabase)
        Me.Name = "frmMain"
        Me.Text = "frmMain"
        CType(Me.dgvDatabase, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabDatabase.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgvDatabase As DataGridView
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents tabDatabase As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents cmbDatabase As ComboBox
    Friend WithEvents lblDatabase As Label
    Friend WithEvents btnCreateBackup As Button
End Class
