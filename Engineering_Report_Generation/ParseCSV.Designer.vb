<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ParseCSV
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
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New DataVisualization.Charting.Series()
        BtnLoadCSV = New Button()
        btnBrowseSaveDir = New Button()
        BtnResampleAndSave = New Button()
        FilesLoaded = New TextBox()
        txtSaveDirectory = New Label()
        Samp = New TextBox()
        Label1 = New Label()
        txtAdditionalPoints = New TextBox()
        Label2 = New Label()
        DataIn = New DataGridView()
        NumCol = New TextBox()
        NumRow = New TextBox()
        lblRowCnt = New Label()
        lblColCnt = New Label()
        lblRcntB = New TextBox()
        lblCcntB = New TextBox()
        Label3 = New Label()
        Label4 = New Label()
        Chart1 = New DataVisualization.Charting.Chart()
        CType(DataIn, ComponentModel.ISupportInitialize).BeginInit()
        CType(Chart1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' BtnLoadCSV
        ' 
        BtnLoadCSV.Location = New Point(12, 12)
        BtnLoadCSV.Name = "BtnLoadCSV"
        BtnLoadCSV.Size = New Size(105, 41)
        BtnLoadCSV.TabIndex = 0
        BtnLoadCSV.Text = "Browse"
        BtnLoadCSV.UseVisualStyleBackColor = True
        ' 
        ' btnBrowseSaveDir
        ' 
        btnBrowseSaveDir.Location = New Point(12, 59)
        btnBrowseSaveDir.Name = "btnBrowseSaveDir"
        btnBrowseSaveDir.Size = New Size(105, 41)
        btnBrowseSaveDir.TabIndex = 1
        btnBrowseSaveDir.Text = "Save Directory"
        btnBrowseSaveDir.UseVisualStyleBackColor = True
        ' 
        ' BtnResampleAndSave
        ' 
        BtnResampleAndSave.Location = New Point(12, 106)
        BtnResampleAndSave.Name = "BtnResampleAndSave"
        BtnResampleAndSave.Size = New Size(105, 41)
        BtnResampleAndSave.TabIndex = 2
        BtnResampleAndSave.Text = "Process Data"
        BtnResampleAndSave.UseVisualStyleBackColor = True
        ' 
        ' FilesLoaded
        ' 
        FilesLoaded.Location = New Point(630, 12)
        FilesLoaded.Multiline = True
        FilesLoaded.Name = "FilesLoaded"
        FilesLoaded.Size = New Size(476, 472)
        FilesLoaded.TabIndex = 3
        ' 
        ' txtSaveDirectory
        ' 
        txtSaveDirectory.AutoSize = True
        txtSaveDirectory.Location = New Point(123, 15)
        txtSaveDirectory.Name = "txtSaveDirectory"
        txtSaveDirectory.Size = New Size(58, 15)
        txtSaveDirectory.TabIndex = 4
        txtSaveDirectory.Text = "Save Path"
        ' 
        ' Samp
        ' 
        Samp.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Samp.Location = New Point(524, 432)
        Samp.Name = "Samp"
        Samp.Size = New Size(100, 23)
        Samp.TabIndex = 5
        Samp.TextAlign = HorizontalAlignment.Center
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(414, 435)
        Label1.Name = "Label1"
        Label1.Size = New Size(104, 15)
        Label1.TabIndex = 6
        Label1.Text = "Sample Interval (s)"
        ' 
        ' txtAdditionalPoints
        ' 
        txtAdditionalPoints.Location = New Point(524, 461)
        txtAdditionalPoints.Name = "txtAdditionalPoints"
        txtAdditionalPoints.Size = New Size(100, 23)
        txtAdditionalPoints.TabIndex = 7
        txtAdditionalPoints.TextAlign = HorizontalAlignment.Center
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(346, 464)
        Label2.Name = "Label2"
        Label2.Size = New Size(172, 15)
        Label2.TabIndex = 8
        Label2.Text = "Number of Samples to Average"
        ' 
        ' DataIn
        ' 
        DataIn.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataIn.Location = New Point(12, 490)
        DataIn.Name = "DataIn"
        DataIn.Size = New Size(1094, 129)
        DataIn.TabIndex = 9
        ' 
        ' NumCol
        ' 
        NumCol.Location = New Point(368, 376)
        NumCol.Name = "NumCol"
        NumCol.Size = New Size(100, 23)
        NumCol.TabIndex = 10
        NumCol.TextAlign = HorizontalAlignment.Center
        ' 
        ' NumRow
        ' 
        NumRow.Location = New Point(368, 347)
        NumRow.Name = "NumRow"
        NumRow.Size = New Size(100, 23)
        NumRow.TabIndex = 11
        NumRow.TextAlign = HorizontalAlignment.Center
        ' 
        ' lblRowCnt
        ' 
        lblRowCnt.AutoSize = True
        lblRowCnt.Location = New Point(479, 350)
        lblRowCnt.Name = "lblRowCnt"
        lblRowCnt.Size = New Size(125, 15)
        lblRowCnt.TabIndex = 13
        lblRowCnt.Text = "Number of Rows After"
        ' 
        ' lblColCnt
        ' 
        lblColCnt.AutoSize = True
        lblColCnt.Location = New Point(479, 379)
        lblColCnt.Name = "lblColCnt"
        lblColCnt.Size = New Size(145, 15)
        lblColCnt.TabIndex = 14
        lblColCnt.Text = "Number of Columns After"
        ' 
        ' lblRcntB
        ' 
        lblRcntB.Location = New Point(249, 347)
        lblRcntB.Name = "lblRcntB"
        lblRcntB.Size = New Size(100, 23)
        lblRcntB.TabIndex = 15
        lblRcntB.TextAlign = HorizontalAlignment.Center
        ' 
        ' lblCcntB
        ' 
        lblCcntB.Location = New Point(249, 376)
        lblCcntB.Name = "lblCcntB"
        lblCcntB.Size = New Size(100, 23)
        lblCcntB.TabIndex = 16
        lblCcntB.TextAlign = HorizontalAlignment.Center
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(110, 350)
        Label3.Name = "Label3"
        Label3.Size = New Size(133, 15)
        Label3.TabIndex = 17
        Label3.Text = "Number of Rows Before"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(90, 379)
        Label4.Name = "Label4"
        Label4.Size = New Size(153, 15)
        Label4.TabIndex = 18
        Label4.Text = "Number of Columns Before"
        ' 
        ' Chart1
        ' 
        ChartArea1.Name = "ChartArea1"
        Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Chart1.Legends.Add(Legend1)
        Chart1.Location = New Point(123, 12)
        Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Chart1.Series.Add(Series1)
        Chart1.Size = New Size(501, 300)
        Chart1.TabIndex = 19
        Chart1.Text = "Chart1"
        ' 
        ' ParseCSV
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1118, 631)
        Controls.Add(Chart1)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(lblCcntB)
        Controls.Add(lblRcntB)
        Controls.Add(lblColCnt)
        Controls.Add(lblRowCnt)
        Controls.Add(NumRow)
        Controls.Add(NumCol)
        Controls.Add(DataIn)
        Controls.Add(Label2)
        Controls.Add(txtAdditionalPoints)
        Controls.Add(Label1)
        Controls.Add(Samp)
        Controls.Add(txtSaveDirectory)
        Controls.Add(FilesLoaded)
        Controls.Add(BtnResampleAndSave)
        Controls.Add(btnBrowseSaveDir)
        Controls.Add(BtnLoadCSV)
        Name = "ParseCSV"
        Text = "Form1"
        CType(DataIn, ComponentModel.ISupportInitialize).EndInit()
        CType(Chart1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents BtnLoadCSV As Button
    Friend WithEvents btnBrowseSaveDir As Button
    Friend WithEvents BtnResampleAndSave As Button
    Friend WithEvents FilesLoaded As TextBox
    Friend WithEvents txtSaveDirectory As Label
    Friend WithEvents Samp As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtAdditionalPoints As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents DataIn As DataGridView
    Friend WithEvents NumCol As TextBox
    Friend WithEvents NumRow As TextBox
    Friend WithEvents lblRowCnt As Label
    Friend WithEvents lblColCnt As Label
    Friend WithEvents lblRcntB As TextBox
    Friend WithEvents lblCcntB As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
End Class
