<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoadDataForReports
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
        MenuStrip1 = New MenuStrip()
        AddRowToolStripMenuItem = New ToolStripMenuItem()
        AddRowToolStripMenuItem1 = New ToolStripMenuItem()
        AddColumnToolStripMenuItem = New ToolStripMenuItem()
        DeleteRowToolStripMenuItem = New ToolStripMenuItem()
        DeleteColumnToolStripMenuItem = New ToolStripMenuItem()
        SaveToolStripMenuItem = New ToolStripMenuItem()
        LoadTablesToolStripMenuItem = New ToolStripMenuItem()
        MonthCalendar1 = New MonthCalendar()
        Grid1 = New DataGridView()
        Label1 = New Label()
        Grid2 = New DataGridView()
        Label2 = New Label()
        Button1 = New Button()
        TextBox1 = New TextBox()
        MenuStrip1.SuspendLayout()
        CType(Grid1, ComponentModel.ISupportInitialize).BeginInit()
        CType(Grid2, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {AddRowToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(1289, 24)
        MenuStrip1.TabIndex = 0
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' AddRowToolStripMenuItem
        ' 
        AddRowToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {AddRowToolStripMenuItem1, AddColumnToolStripMenuItem, DeleteRowToolStripMenuItem, DeleteColumnToolStripMenuItem, SaveToolStripMenuItem, LoadTablesToolStripMenuItem})
        AddRowToolStripMenuItem.Name = "AddRowToolStripMenuItem"
        AddRowToolStripMenuItem.Size = New Size(37, 20)
        AddRowToolStripMenuItem.Text = "File"
        ' 
        ' AddRowToolStripMenuItem1
        ' 
        AddRowToolStripMenuItem1.Name = "AddRowToolStripMenuItem1"
        AddRowToolStripMenuItem1.ShortcutKeys = Keys.Control Or Keys.R
        AddRowToolStripMenuItem1.Size = New Size(184, 22)
        AddRowToolStripMenuItem1.Text = "Add Row"
        ' 
        ' AddColumnToolStripMenuItem
        ' 
        AddColumnToolStripMenuItem.Name = "AddColumnToolStripMenuItem"
        AddColumnToolStripMenuItem.ShortcutKeys = Keys.Control Or Keys.C
        AddColumnToolStripMenuItem.Size = New Size(184, 22)
        AddColumnToolStripMenuItem.Text = "Add Column"
        ' 
        ' DeleteRowToolStripMenuItem
        ' 
        DeleteRowToolStripMenuItem.Name = "DeleteRowToolStripMenuItem"
        DeleteRowToolStripMenuItem.Size = New Size(184, 22)
        DeleteRowToolStripMenuItem.Text = "Delete Row"
        ' 
        ' DeleteColumnToolStripMenuItem
        ' 
        DeleteColumnToolStripMenuItem.Name = "DeleteColumnToolStripMenuItem"
        DeleteColumnToolStripMenuItem.Size = New Size(184, 22)
        DeleteColumnToolStripMenuItem.Text = "Delete Column"
        ' 
        ' SaveToolStripMenuItem
        ' 
        SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        SaveToolStripMenuItem.ShortcutKeys = Keys.Control Or Keys.S
        SaveToolStripMenuItem.Size = New Size(184, 22)
        SaveToolStripMenuItem.Text = "Save"
        ' 
        ' LoadTablesToolStripMenuItem
        ' 
        LoadTablesToolStripMenuItem.Name = "LoadTablesToolStripMenuItem"
        LoadTablesToolStripMenuItem.Size = New Size(184, 22)
        LoadTablesToolStripMenuItem.Text = "Load Tables"
        ' 
        ' MonthCalendar1
        ' 
        MonthCalendar1.Location = New Point(1044, 64)
        MonthCalendar1.Name = "MonthCalendar1"
        MonthCalendar1.TabIndex = 1
        ' 
        ' Grid1
        ' 
        Grid1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        Grid1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllHeaders
        Grid1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Grid1.Location = New Point(12, 64)
        Grid1.Name = "Grid1"
        Grid1.Size = New Size(1020, 205)
        Grid1.TabIndex = 2
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 46)
        Label1.Name = "Label1"
        Label1.Size = New Size(43, 15)
        Label1.TabIndex = 3
        Label1.Text = "Table 1"
        ' 
        ' Grid2
        ' 
        Grid2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Grid2.Location = New Point(12, 298)
        Grid2.Name = "Grid2"
        Grid2.Size = New Size(1020, 201)
        Grid2.TabIndex = 4
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(12, 280)
        Label2.Name = "Label2"
        Label2.Size = New Size(43, 15)
        Label2.TabIndex = 5
        Label2.Text = "Table 2"
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(1151, 238)
        Button1.Name = "Button1"
        Button1.Size = New Size(120, 23)
        Button1.TabIndex = 6
        Button1.Text = "Export Tables"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(1053, 298)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(218, 23)
        TextBox1.TabIndex = 7
        ' 
        ' Form2
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1289, 537)
        Controls.Add(TextBox1)
        Controls.Add(Button1)
        Controls.Add(Label2)
        Controls.Add(Grid2)
        Controls.Add(Label1)
        Controls.Add(Grid1)
        Controls.Add(MonthCalendar1)
        Controls.Add(MenuStrip1)
        MainMenuStrip = MenuStrip1
        Name = "Form2"
        Text = "Form2"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        CType(Grid1, ComponentModel.ISupportInitialize).EndInit()
        CType(Grid2, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents AddRowToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AddRowToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents AddColumnToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DeleteRowToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DeleteColumnToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MonthCalendar1 As MonthCalendar
    Friend WithEvents Grid1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents Grid2 As DataGridView
    Friend WithEvents Label2 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents LoadTablesToolStripMenuItem As ToolStripMenuItem
End Class
