Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.FileIO
Imports ClosedXML
Imports ClosedXML.Excel
Imports ClosedXML.Excel.Exceptions
Imports DocumentFormat.OpenXml.Wordprocessing


Public Class ParseCSV

    Private DataTbSave As New DataTable

    Private DataGridViewList As New List(Of DataGridView)()

    Private Sub BtnLoadCSV_Click(sender As Object, e As EventArgs) Handles BtnLoadCSV.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        openFileDialog.Multiselect = False ' Load one file at a time


        If openFileDialog.ShowDialog() = DialogResult.OK Then
            LoadCSVToNewDataGridView(openFileDialog.FileName)
        End If


    End Sub

    Private Sub LoadCSVToNewDataGridView(filePath As String)
        ' Create new DataGridView
        Dim dgv As New DataGridView()
        dgv.Width = 500
        dgv.Height = 200
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.Dock = DockStyle.Top

        ' Read CSV and populate DataGridView
        Dim dt As New DataTable()
        Dim lines() As String = File.ReadAllLines(filePath)
        FilesLoaded.Text = filePath


        If lines.Length > 0 Then

            ' Read first two rows and combine them as headers
            Dim headers1() As String = lines(0).Split(","c)
            Dim headers2() As String = lines(1).Split(","c)

            Dim combinedHeaders(headers1.Length - 1) As String

            For i As Integer = 0 To headers1.Length - 1
                combinedHeaders(i) = headers1(i).Trim() & " " & headers2(i).Trim()
                dt.Columns.Add(combinedHeaders(i))
            Next

            ' Format decimal numbers to 3 decimal places
            AddHandler dgv.CellFormatting, AddressOf FormatDecimalCells

            ' Add data rows
            For i As Integer = 2 To lines.Length - 1
                dt.Rows.Add(lines(i).Split(","c))
            Next

        End If

        dgv.DataSource = dt

        dgv.Visible = True
        DataGridViewList.Add(dgv)

        ' Add DataGridView to Form
        Me.Controls.Add(dgv)


    End Sub

    Private Sub FormatDecimalCells(sender As Object, e As DataGridViewCellFormattingEventArgs)
        Dim dgv As DataGridView = DirectCast(sender, DataGridView)

        ' Check if the cell contains a numeric value and format it
        If e.Value IsNot Nothing AndAlso IsNumeric(e.Value) Then
            Dim num As Double
            If Double.TryParse(e.Value.ToString(), num) Then
                e.Value = num.ToString("F3") ' Display with 3 decimal places
                e.FormattingApplied = True ' Prevents infinite recursion
            End If
        End If

    End Sub

    Private Sub BtnResampleAndSave_Click(sender As Object, e As EventArgs) Handles BtnResampleAndSave.Click
        If DataGridViewList.Count > 0 Then
            Dim dgv As DataGridView = DataGridViewList(0) ' Get the first DataGridView



            ' Parameters for resampling and averaging
            Dim sampleInterval As Double = CDbl(Samp.Text) ' Interval in time or index units (adjust as needed)
            Dim numSamplesToAverage As Integer = CInt(txtAdditionalPoints.Text) ' Number of samples to average around each resample point



            ' Resample the data
            Dim resampledData As DataTable = ResampleData(dgv, sampleInterval, numSamplesToAverage)

            ' Create new DataGridView for the resampled data
            Dim newDgv As New DataGridView()

            newDgv.Width = 800
            newDgv.Height = 250
            newDgv.ReadOnly = True
            newDgv.AllowUserToAddRows = False
            newDgv.Dock = DockStyle.Top
            newDgv.DataSource = resampledData



            ' Add the new DataGridView to the form and the list
            Me.Controls.Add(newDgv)
            DataGridViewList.Add(newDgv)
            newDgv.Visible = False



            NumRow.Text = newDgv.RowCount.ToString
            NumCol.Text = newDgv.ColumnCount.ToString
            DataTbSave = resampledData


        Else
            MessageBox.Show("No data to process.")
        End If
    End Sub

    Private Function ResampleData(dgv As DataGridView, sampleInterval As Double, numSamplesToAverage As Integer) As DataTable
        Dim dt As DataTable = CType(dgv.DataSource, DataTable)
        Dim resampledData As New DataTable()

        ' Copy the column headers
        For Each col As DataColumn In dt.Columns
            resampledData.Columns.Add(col.ColumnName)
        Next

        ' Iterate through the rows and resample
        Dim resampledRows As New List(Of DataRow)()
        Dim rowCount As Integer = dt.Rows.Count

        For i As Integer = 0 To rowCount - 1 Step CInt(sampleInterval)
            Dim startIndex As Integer = Math.Max(i - CInt(numSamplesToAverage / 2), 0)
            Dim endIndex As Integer = Math.Min(i + CInt(numSamplesToAverage / 2), rowCount - 1)

            ' Get the data for averaging
            Dim averagedRow As DataRow = resampledData.NewRow()

            For j As Integer = 0 To dt.Columns.Count - 1
                ' If the column contains numeric data, average it
                If IsNumeric(dt.Rows(i)(j)) Then
                    Dim sum As Double = 0
                    Dim count As Integer = 0
                    For k As Integer = startIndex To endIndex
                        If IsNumeric(dt.Rows(k)(j)) Then
                            sum += Convert.ToDouble(dt.Rows(k)(j))
                            count += 1
                        End If
                    Next
                    averagedRow(j) = sum / count ' Average the values
                Else
                    'averagedRow(j) = dt.Rows(i)(j) ' Copy non-numeric data as is

                End If
            Next

            resampledRows.Add(averagedRow)
        Next

        ' Add the resampled rows to the DataTable
        For Each row In resampledRows
            resampledData.Rows.Add(row)
        Next

        Return resampledData
    End Function

    Private Sub SaveDataToCSV(resampledData As DataTable)
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        saveFileDialog.DefaultExt = "csv"
        saveFileDialog.FileName = "Parse_"

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Using writer As New StreamWriter(saveFileDialog.FileName)
                ' Write the header
                Dim columnNames As String = String.Join(",", resampledData.Columns.Cast(Of DataColumn)().Select(Function(c) c.ColumnName))
                writer.WriteLine(columnNames)

                ' Write the data
                For Each row As DataRow In resampledData.Rows
                    Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(field) field.ToString()))
                    writer.WriteLine(rowData)
                Next
            End Using
            MessageBox.Show("Data saved successfully.")
        End If
    End Sub

    Private Sub btnBrowseSaveDir_Click(sender As Object, e As EventArgs) Handles btnBrowseSaveDir.Click
        SaveDataToCSV(DataTbSave)
    End Sub

End Class

