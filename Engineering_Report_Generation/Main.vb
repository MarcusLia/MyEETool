Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Data
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Text
Imports Microsoft.Office.Interop



Public Class Main
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox1.Text, TextBox1.Font)
        TextBox1.Width = Math.Min(300, txtSize.Width + 10)
        TextBox1.TextAlign = HorizontalAlignment.Center
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox2.Text, TextBox2.Font)
        TextBox2.Width = Math.Min(300, txtSize.Width + 10)
        TextBox2.TextAlign = HorizontalAlignment.Center
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox3.Text, TextBox3.Font)
        TextBox3.Width = Math.Min(300, txtSize.Width + 10)
        TextBox3.TextAlign = HorizontalAlignment.Center
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox4.Text, TextBox4.Font)
        TextBox4.Width = Math.Min(300, txtSize.Width + 10)
        TextBox4.TextAlign = HorizontalAlignment.Center
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox5.Text, TextBox5.Font)
        TextBox5.Width = Math.Min(300, txtSize.Width + 10)
        TextBox5.TextAlign = HorizontalAlignment.Center
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox6.Text, TextBox6.Font)
        TextBox6.Width = Math.Min(300, txtSize.Width + 10)
        TextBox6.TextAlign = HorizontalAlignment.Center
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox7.Text, TextBox7.Font)
        TextBox7.Width = Math.Min(300, txtSize.Width + 10)
        TextBox7.TextAlign = HorizontalAlignment.Center
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox8.Text, TextBox8.Font)
        TextBox8.Width = Math.Min(300, txtSize.Width + 10)
        TextBox8.TextAlign = HorizontalAlignment.Center
    End Sub


    Private Sub Generate_Click(sender As Object, e As EventArgs) Handles Generate.Click
        Dim templatePath As String = "C:\Template\Engineering_Test_Report_Template.docx"
        Dim savePath As String = "C:\Temp\PowerRunReport.docx"
        Dim chart1 As System.Windows.Forms.DataVisualization.Charting.Chart

        ' Open Word application
        Dim wordApp As New Application()
        Dim wordDoc As Document = wordApp.Documents.Open(templatePath)

        ' Make Word visible (optional)
        wordApp.Visible = True


        InsertTextIntoBookmark(wordDoc, "Title", "McLaren Performance Engineering - Engine Performance")
        InsertTextIntoBookmark(wordDoc, "Address", "32233 West Eight Mile Rd., Livonia MI 48150")
        InsertTextIntoBookmark(wordDoc, "Customer", TextBox1.Text)
        InsertTextIntoBookmark(wordDoc, "EngineType", TextBox2.Text)
        InsertTextIntoBookmark(wordDoc, "TestType", TextBox3.Text)
        InsertTextIntoBookmark(wordDoc, "Oil", TextBox4.Text)
        InsertTextIntoBookmark(wordDoc, "Camshaft", TextBox5.Text)
        InsertTextIntoBookmark(wordDoc, "CamshaftPosition", TextBox6.Text)
        InsertTextIntoBookmark(wordDoc, "ExhaustPrimary", TextBox7.Text)
        InsertTextIntoBookmark(wordDoc, "ExhaustSecondary", TextBox8.Text)
        InsertTextIntoBookmark(wordDoc, "Manifold", TextBox9.Text)
        InsertTextIntoBookmark(wordDoc, "Muffler", TextBox10.Text)

        InsertTextIntoBookmark(wordDoc, "Date", TextBox11.Text)
        InsertTextIntoBookmark(wordDoc, "EngineSerial", TextBox12.Text)
        InsertTextIntoBookmark(wordDoc, "Initials", TextBox13.Text)
        InsertTextIntoBookmark(wordDoc, "FileNo", TextBox14.Text)
        InsertTextIntoBookmark(wordDoc, "InletAir", TextBox15.Text)
        InsertTextIntoBookmark(wordDoc, "DataCorrection", TextBox16.Text)
        InsertTextIntoBookmark(wordDoc, "Displacement", TextBox17.Text)
        InsertTextIntoBookmark(wordDoc, "FiringOrder", TextBox18.Text)
        InsertTextIntoBookmark(wordDoc, "Fuel", TextBox19.Text)
        InsertTextIntoBookmark(wordDoc, "CompressionRatio", TextBox20.Text)
        InsertTextIntoBookmark(wordDoc, "DynoNo", TextBox21.Text)
        InsertTextIntoBookmark(wordDoc, "Notes", TextBox22.Text)

        ExportToWordTemplate(DataGrid1, DataGrid2, wordDoc)

        Dim range As Word.Range = wordDoc.Range()
        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        range.InsertParagraphAfter()

        ExportChartToWord(Chart, "", wordDoc)


        wordDoc.SaveAs2(savePath)
        'wordDoc.Close()
        'wordApp.Quit()

        MessageBox.Show("Report saved successfully at " & savePath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ' Function to insert text into a bookmark
    Private Sub InsertTextIntoBookmark(doc As Document, bookmarkName As String, text As String)
        If doc.Bookmarks.Exists(bookmarkName) Then
            Dim bmRange As Range = doc.Bookmarks(bookmarkName).Range
            bmRange.Text = text
            doc.Bookmarks.Add(bookmarkName, bmRange) ' Re-add the bookmark after inserting text
        End If
    End Sub

    ' Button Click Event to Open CSV File
    Private Sub LinkDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LinkDataToolStripMenuItem.Click
        ' Open File Dialog to Select CSV File
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        openFileDialog.Title = "Select a CSV File"

        ' If user selects a file
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            LoadCSVIntoDataGridView(filePath, DataMisc) ' Call function to load CSV
        End If
    End Sub

    ' Function to Load CSV Data into DataGridView
    Private Sub LoadCSVIntoDataGridView(filePath As String, dataGridView As DataGridView)
        Dim dt As New System.Data.DataTable()
        Dim dataStartFound As Boolean = False
        Dim headerStartFound As Boolean = False

        Try
            ' Clear existing data
            dataGridView.DataSource = Nothing
            dataGridView.Rows.Clear()
            dataGridView.Columns.Clear()
            ListBox1.Items.Clear()



            ' Read CSV File
            Using reader As New StreamReader(filePath)
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    Dim dataStartline() As String = line.Split(","c)
                    ' Populate Header List Box
                    If dataStartline(0) = "[HEADER START]" Then
                        headerStartFound = True

                    End If

                    If headerStartFound Then
                        ListBox1.Items.Add(line)
                    End If

                    If dataStartline(0) = "[HEADER END]" Then
                        headerStartFound = False
                    End If

                    ' Wait until [Data Start] is found
                    If dataStartline(0) = "[DATA START]" Then
                        dataStartFound = True
                        Continue While ' Skip to the next line (header line)
                    End If

                    ' If [Data Start] was found, read the next line as headers
                    If dataStartFound AndAlso dt.Columns.Count = 0 Then
                        Dim headers() As String = line.Split(","c)
                        For Each header As String In headers
                            dt.Columns.Add(header.Trim()) ' Add headers to DataTable
                        Next
                        Continue While ' Skip to the next line (start reading data)
                    End If

                    ' If headers are set, start reading data
                    If dataStartFound Then
                        Dim rows() As String = line.Split(","c)
                        dt.Rows.Add(rows)
                    End If
                End While
            End Using

            ' Bind DataTable to DataGridView
            dataGridView.DataSource = dt

            ' Auto-size the columns
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            PopulateColumnSelector()


        Catch ex As Exception
            MessageBox.Show("Error loading CSV: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Call this method after loading the CSV into DataGridView
    Private Sub PopulateColumnSelector()
        PlotPicker.Items.Clear()

        ' Add column names to ComboBox
        For Each col As DataGridViewColumn In DataMisc.Columns
            PlotPicker.Items.Add(col.HeaderText)
        Next

        ' Select the first column by default
        If PlotPicker.Items.Count > 0 Then
            PlotPicker.SelectedIndex = 0
        End If
    End Sub

    Private Sub PlotPicker_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlotPicker.SelectedIndexChanged
        Chart.ChartAreas(0).AxisX.Minimum = Double.NaN ' Auto Min
        Chart.ChartAreas(0).AxisX.Maximum = Double.NaN ' Auto Max
        Chart.ChartAreas(0).AxisX.Interval = Double.NaN ' Auto Interval

        Chart.ChartAreas(0).AxisY.Minimum = Double.NaN ' Auto Min
        Chart.ChartAreas(0).AxisY.Maximum = Double.NaN ' Auto Max
        Chart.ChartAreas(0).AxisY.Interval = Double.NaN ' Auto Interval

        PlotSelectedColumn()
        Chart.ChartAreas(0).RecalculateAxesScale() ' Force re-scale
    End Sub
    Private Sub PlotSelectedColumn()
        ' Ensure a column is selected
        If PlotPicker.SelectedIndex = -1 Then
            MessageBox.Show("Please select a column to display.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Get selected column name
        Dim selectedColumn As String = PlotPicker.SelectedItem.ToString()

        ' Clear existing chart data
        ' Chart.Series.Clear()

        ' Create new series for the selected column
        Dim series As New System.Windows.Forms.DataVisualization.Charting.Series(selectedColumn)

        With Chart.ChartAreas(0)
            ' X-Axis Configuration
            .AxisX.Title = "Time (s)" ' Set X-Axis Title
            ' Y-Axis Configuration
            .AxisY.Title = selectedColumn ' Set Y-Axis Title

        End With

        series.ChartType = SeriesChartType.Line ' Change to Bar, Column, etc. if needed
        series.IsVisibleInLegend = True ' Ensure the series appears in the legend
        ' Loop through DataGridView rows and add data points
        For Each row As DataGridViewRow In DataMisc.Rows
            If Not row.IsNewRow Then
                Dim xValue As Object = row.Cells(0).Value ' Assuming first column is X-axis (Speed)
                Dim yValue As Object = row.Cells(selectedColumn).Value ' Selected column as Y-axis

                ' Ensure values are numeric before adding
                If IsNumeric(xValue) AndAlso IsNumeric(yValue) Then
                    series.Points.AddXY(Convert.ToDouble(xValue), Convert.ToDouble(yValue))
                End If
            End If
        Next

        ' Add series to the chart
        Chart.Series.Add(series)

        With Chart.ChartAreas(0)
            .CursorX.IsUserEnabled = True
            .CursorX.IsUserSelectionEnabled = True
            .AxisX.ScaleView.Zoomable = True
        End With

        ' Enable and customize the X-axis cursor
        Chart.ChartAreas(0).CursorX.IsUserEnabled = True
        Chart.ChartAreas(0).CursorX.IsUserSelectionEnabled = True
        Chart.ChartAreas(0).CursorX.LineColor = Color.Red
        Chart.ChartAreas(0).CursorX.LineWidth = 1
        Chart.ChartAreas(0).CursorX.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash

        ' Enable and customize the Y-axis cursor
        Chart.ChartAreas(0).CursorY.IsUserEnabled = True
        Chart.ChartAreas(0).CursorY.IsUserSelectionEnabled = True
        Chart.ChartAreas(0).CursorY.LineColor = Color.Blue
        Chart.ChartAreas(0).CursorY.LineWidth = 1
        Chart.ChartAreas(0).CursorY.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash

    End Sub

    Private Sub Chart1_AxisViewChanged(sender As Object, e As ViewEventArgs) Handles Chart.AxisViewChanged
        Dim xMin As Double = e.Axis.ScaleView.ViewMinimum
        Dim xMax As Double = e.Axis.ScaleView.ViewMaximum

        ' Initialize DataGridView if needed
        If DataGrid1.Columns.Count = 0 Then
            DataGrid1.Columns.Add("SeriesName", "Series Name")
            DataGrid1.Columns.Add("MinValue", "Min")
            DataGrid1.Columns.Add("MaxValue", "Max")
            DataGrid1.Columns.Add("AvgValue", "Avg")
        End If

        DataGrid1.Rows.Clear() ' Clear old data

        ' Loop through each series and compute stats
        For Each series As System.Windows.Forms.DataVisualization.Charting.Series In Chart.Series
            Dim filteredPoints = series.Points.Where(Function(p) p.XValue >= xMin AndAlso p.XValue <= xMax).ToList()

            If filteredPoints.Count > 0 Then
                Dim values As List(Of Double) = filteredPoints.Select(Function(p) p.YValues(0)).ToList()

                ' Compute statistics
                Dim minVal As Double = values.Min()
                Dim maxVal As Double = values.Max()
                Dim avgVal As Double = values.Average()

                ' Add statistics to DataGridView
                DataGrid1.Rows.Add(series.Name, minVal, maxVal, avgVal)
            End If
        Next

    End Sub



    Private Sub Chart_CursorPositionChanged(sender As Object, e As System.Windows.Forms.DataVisualization.Charting.CursorEventArgs) Handles Chart.CursorPositionChanged
        Dim xPos As Double = e.NewPosition
        ' MessageBox.Show("Cursor X Position: " & xPos.ToString())
    End Sub

    Private Sub Chart_MouseMove(sender As Object, e As MouseEventArgs) Handles Chart.MouseMove
        Dim hit As HitTestResult = Chart.HitTest(e.X, e.Y)

        ' Check if the mouse is over a data point
        If hit.ChartElementType = ChartElementType.DataPoint Then
            Dim seriesIndex As Integer = hit.Series.Points.IndexOf(hit.Object)
            Dim xValue As Double = hit.Series.Points(seriesIndex).XValue
            Dim yValue As Double = hit.Series.Points(seriesIndex).YValues(0)
            Dim seriesName As String = hit.Series.Name
            ' Snap cursor to this data point
            Chart.ChartAreas(0).CursorX.Position = xValue
            Chart.ChartAreas(0).CursorY.Position = yValue

            ' Show coordinates in a label
            lblCoordinates.Text = $"X: {xValue:F2}, Y: {yValue:F2}"
            lblSelectedPlot.Text = "Selected Plot: " & seriesName '
        End If
    End Sub

    Private Sub btnResetZoom_Click(sender As Object, e As EventArgs) Handles btnResetZoom.Click
        With Chart.ChartAreas(0).AxisX
            .ScaleView.ZoomReset() ' Reset X-Axis zoom
        End With

        With Chart.ChartAreas(0).AxisY
            .ScaleView.ZoomReset() ' Reset Y-Axis zoom
        End With
        Chart.Series.Clear()

    End Sub

    Private Sub Restore_Click(sender As Object, e As EventArgs) Handles Restore.Click
        With Chart.ChartAreas(0).AxisX
            .ScaleView.ZoomReset() ' Reset X-Axis zoom
        End With

        With Chart.ChartAreas(0).AxisY
            .ScaleView.ZoomReset() ' Reset Y-Axis zoom
        End With
    End Sub

    Private Sub AddToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddToolStripMenuItem.Click
        LoadDataForReports.Visible = True
    End Sub


    Public Sub CopyDataFromGrid(ByVal dgv As DataGridView, ByVal dgv2 As DataGridView)
        DataGrid1.Rows.Clear()
        DataGrid1.Columns.Clear()
        DataGrid2.Rows.Clear()
        DataGrid2.Columns.Clear()


        ' Copy column headers
        For Each col As DataGridViewColumn In dgv.Columns
            DataGrid1.Columns.Add(col.Clone())
        Next

        For Each col As DataGridViewColumn In dgv2.Columns
            DataGrid2.Columns.Add(col.Clone())
        Next

        ' Copy row data
        For Each row As DataGridViewRow In dgv.Rows
            If Not row.IsNewRow Then
                Dim newRow As DataGridViewRow = CType(row.Clone(), DataGridViewRow)
                For colIndex As Integer = 0 To row.Cells.Count - 1
                    newRow.Cells(colIndex).Value = row.Cells(colIndex).Value
                Next
                DataGrid1.Rows.Add(newRow)
            End If
        Next

        For Each row As DataGridViewRow In dgv2.Rows
            If Not row.IsNewRow Then
                Dim newRow As DataGridViewRow = CType(row.Clone(), DataGridViewRow)
                For colIndex As Integer = 0 To row.Cells.Count - 1
                    newRow.Cells(colIndex).Value = row.Cells(colIndex).Value
                Next
                DataGrid2.Rows.Add(newRow)
            End If
        Next

    End Sub

    Public Sub ExportToWordTemplate(ByVal dgv As DataGridView, ByVal dgv2 As DataGridView, ByVal wordDoc As Word.Document)

        ' Ensure the document has at least one table
        If wordDoc.Tables.Count = 0 Then
            MessageBox.Show("No table found in the document.")
            wordDoc.Close(False)

            Exit Sub
        End If

        ' Get the Misc Data table in the document
        Dim range As Word.Range = wordDoc.Range()
        'Dim table As Word.Table = wordDoc.Tables(6)
        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        range.InsertParagraphAfter()

        ' Create a table below the existing content
        Dim rowCount As Integer = dgv.Rows.Count
        Dim colCount As Integer = dgv.Columns.Count
        Dim table As Word.Table = wordDoc.Tables.Add(range, rowCount + 1, colCount)


        For colIndex As Integer = 0 To dgv.ColumnCount - 1
            table.Cell(1, colIndex + 1).Range.Text = dgv.Columns(colIndex).HeaderText
            table.Cell(1, colIndex + 1).Range.Font.Bold = 1  ' Make headers bold
        Next

        ' Ensure the table has enough rows
        For rowIndex As Integer = 0 To dgv.RowCount - 1
            ' If table does not have enough rows, add a new row
            If rowIndex + 2 > table.Rows.Count Then
                table.Rows.Add()
            End If

            For colIndex As Integer = 0 To dgv.ColumnCount - 1

                If colIndex + 1 > table.Columns.Count Then
                    'table.Columns.Add()
                End If
                ' Ensure column index is within the table's column count
                If colIndex + 1 <= table.Columns.Count Then
                    table.Columns.Add()
                    ' Handle null values safely
                    Dim cellValue As String = If(dgv.Rows(rowIndex).Cells(colIndex).Value?.ToString(), "")
                    table.Cell(rowIndex + 2, colIndex + 1).Range.Text = cellValue
                End If
            Next
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim afterTableRange As Word.Range = wordDoc.Range()
        afterTableRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        afterTableRange.InsertParagraphAfter()

        Dim rowCount2 As Integer = dgv2.Rows.Count
        Dim colCount2 As Integer = dgv2.Columns.Count
        Dim table2 As Word.Table = wordDoc.Tables.Add(range, rowCount2 + 1, colCount2)

        For colIndex As Integer = 0 To dgv2.ColumnCount - 1
            table2.Cell(1, colIndex + 1).Range.Text = dgv2.Columns(colIndex).HeaderText
            table2.Cell(1, colIndex + 1).Range.Font.Bold = 1  ' Make headers bold
        Next

        ' Ensure the table has enough rows
        For rowIndex As Integer = 0 To dgv2.RowCount - 1
            ' If table does not have enough rows, add a new row
            If rowIndex + 2 > table.Rows.Count Then
                table2.Rows.Add()
            End If

            For colIndex As Integer = 0 To dgv2.ColumnCount - 1
                ' Ensure column index is within the table's column count
                If colIndex + 1 > table2.Columns.Count Then
                    'table2.Columns.Add()
                End If
                If colIndex + 1 <= table2.Columns.Count Then

                    ' Handle null values safely
                    Dim cellValue As String = If(dgv2.Rows(rowIndex).Cells(colIndex).Value?.ToString(), "")
                    table2.Cell(rowIndex + 2, colIndex + 1).Range.Text = cellValue
                End If
            Next
        Next

        ' Save and close the document

    End Sub



    ' This method exports the chart as an image to a Word document
    Public Sub ExportChartToWord(chart As DataVisualization.Charting.Chart, filePath As String, ByVal wordDoc As Word.Document)

        Dim chartImagePath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "chart.png")
        chart.SaveImage(chartImagePath, Drawing.Imaging.ImageFormat.Png)

        If wordDoc.Bookmarks.Exists("Charts") Then
            ' Get the bookmark range
            Dim bookmarkRange As Word.Range = wordDoc.Bookmarks("Charts").Range



            Dim section As Word.Section = bookmarkRange.Sections.First
            section.PageSetup.TextColumns.SetCount(1)

            ' Insert the image at the bookmark location
            bookmarkRange.InlineShapes.AddPicture(chartImagePath)

            ' Optional: Update bookmark position after inserting the image
            wordDoc.Bookmarks.Add("Charts", bookmarkRange)
        Else
            MessageBox.Show("Bookmark '" & "Charts" & "' not found in document.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Public Sub CreateNewLabel(ByVal lblNew As Label, ByVal dataIn As String)


        ' Set label properties
        lblNew.Text = lblNew.Name + ":" + dataIn
        lblNew.Font = New Drawing.Font("Arial", 12, FontStyle.Bold)
        lblNew.AutoSize = True
        lblNew.Location = New Drawing.Point(10, 10)
        Panel1.Controls.Add(lblNew)
        ' Add label to the form
        ' Me.Controls.Add(newLabel)
        'Number One COmment for GitHub Practice


    End Sub

    Private Sub DataDisplayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataDisplayToolStripMenuItem.Click
        Dim newLabel As New Label
        Dim userInput As String = InputBox("Enter Name of New Data: ", "User Input")
        newLabel.Name = userInput
        CreateNewLabel(newLabel, "135ft-lb")

    End Sub

    Private Sub ParseCSVToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ParseCSVToolStripMenuItem.Click
        ParseCSV.Visible = True

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        Dim txtSize As Size = TextRenderer.MeasureText(TextBox14.Text, TextBox1.Font)
        TextBox14.Width = Math.Min(300, txtSize.Width + 10)
        TextBox14.TextAlign = HorizontalAlignment.Center
    End Sub

    ' Helper function to release Word objects










End Class
