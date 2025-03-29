Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.IO
Imports System.Data
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices




Public Class LoadDataForReports

    Dim timesRowClicked As Integer = 0
    Dim timesColumnClicked As Integer = 0
    Dim largest As Integer = 0
    Dim x As Integer = 0
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Workbook
    Dim xlWorksheet As Worksheet
    Dim xlRange As Range

    Private Sub AddRowToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AddRowToolStripMenuItem1.Click
        If Grid1.Focused = True Then
            Dim userInput As String = InputBox("Enter Row Name:", "User Input")
            Grid1.Rows().Add(userInput)
            Grid1.RowHeadersVisible = True
            Grid1.AutoResizeRows()
            Grid1.AutoResizeColumns()
        End If

        If Grid2.Focused = True Then
            Dim userInput As String = InputBox("Enter Row Name:", "User Input")
            Grid2.Rows().Add(userInput)
            Grid2.RowHeadersVisible = True
            Grid2.AutoResizeRows()
            Grid2.AutoResizeColumns()
        End If

    End Sub

    Private Sub AddColumnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddColumnToolStripMenuItem.Click

        If Grid1.Focused = True Then
            Dim userInput As String = InputBox("Enter Column Name:", "User Input")
            Grid1.Columns().Add(userInput, userInput)
            Grid1.AutoResizeRows()
            Grid1.AutoResizeColumns()
        End If
        If Grid2.Focused = True Then
            Dim userInput As String = InputBox("Enter Column Name:", "User Input")
            Grid2.Columns().Add(userInput, userInput)
            Grid2.AutoResizeRows()
            Grid2.AutoResizeColumns()
        End If

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid1.ColumnCount() = 1
        Grid2.ColumnCount() = 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ExportToMainForm()

    End Sub

    Public Sub ExportToMainForm()

        Dim mainForm As Main = CType(System.Windows.Forms.Application.OpenForms("Form1"), Main)
        mainForm?.CopyDataFromGrid(Grid1, Grid2)

    End Sub

    Private Sub DeleteRowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteRowToolStripMenuItem.Click
        If Grid1.Focused = True Then
            Dim rowSelected As Integer
            rowSelected = Grid1.CurrentRow().Index
            TextBox1.Text = rowSelected
            Grid1.Rows().RemoveAt(rowSelected)
        End If

        If Grid2.Focused = True Then
            Dim rowSelected As Integer
            rowSelected = Grid2.CurrentRow().Index
            TextBox1.Text = rowSelected
            Grid2.Rows().RemoveAt(rowSelected)
        End If
    End Sub

    Private Sub DeleteColumnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteColumnToolStripMenuItem.Click

        If Grid1.Focused = True Then
            Dim columnSelected As Integer
            columnSelected = Grid1.SelectedCells(0).ColumnIndex
            TextBox1.Text = columnSelected
            Grid1.Columns().RemoveAt(columnSelected)
        End If

        If Grid2.Focused = True Then
            Dim columnSelected As Integer
            columnSelected = Grid2.SelectedCells(0).ColumnIndex
            TextBox1.Text = columnSelected
            Grid2.Columns().RemoveAt(columnSelected)
        End If


    End Sub


    Public Sub ExportTwoDataGridViewsToExcel(dgv1 As DataGridView, dgv2 As DataGridView, filePath As String)
        ' Create an Excel application object
        Dim excelApp As New Excel.Application()
        excelApp.Visible = False  ' Set to True if you want to see the Excel window

        ' Create a new workbook and get the active sheet
        Dim workbooks As Excel.Workbooks = excelApp.Workbooks
        Dim workbook As Excel.Workbook = workbooks.Add()

        ' Export the first DataGridView to the first sheet
        Dim worksheet1 As Excel.Worksheet = workbook.Sheets(1)
        ExportDataGridViewToSheet(dgv1, worksheet1)

        ' Create a second sheet for the second DataGridView
        Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=workbook.Sheets(workbook.Sheets.Count))
        worksheet2.Name = "Table2"
        ExportDataGridViewToSheet(dgv2, worksheet2)

        ' Save the Excel file
        workbook.SaveAs(filePath)

        ' Clean up
        workbook.Close()
        excelApp.Quit()

        ' Release Excel objects to free memory
        ReleaseObject(worksheet1)
        ReleaseObject(worksheet2)
        ReleaseObject(workbook)
        ReleaseObject(workbooks)
        ReleaseObject(excelApp)

        MessageBox.Show("Data exported to Excel successfully!")
    End Sub

    ' This helper function writes DataGridView data to a given Excel worksheet
    Private Sub ExportDataGridViewToSheet(dgv As DataGridView, worksheet As Excel.Worksheet)
        ' Add headers to the first row of Excel from DataGridView
        For colIndex As Integer = 0 To dgv.ColumnCount - 1
            worksheet.Cells(1, colIndex + 1) = dgv.Columns(colIndex).HeaderText
        Next

        ' Add DataGridView rows to Excel starting from the second row
        For rowIndex As Integer = 0 To dgv.RowCount - 1
            For colIndex As Integer = 0 To dgv.ColumnCount - 1
                worksheet.Cells(rowIndex + 2, colIndex + 1) = dgv.Rows(rowIndex).Cells(colIndex).Value
            Next
        Next
    End Sub

    ' This helper function releases Excel objects to avoid memory leaks
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click

        ' Open FolderBrowserDialog to allow the user to select the folder
        Dim folderDialog As New FolderBrowserDialog()
        folderDialog.Description = "Select the folder to save the Excel file"

        If folderDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected folder path
            Dim folderPath As String = folderDialog.SelectedPath

            ' Create the full file path for the Excel file (you can specify the file name here)
            Dim filePath As String = System.IO.Path.Combine(folderPath, "Data_Table_1.xlsx")

            ' Call the export function with the DataGridView and file path
            ExportTwoDataGridViewsToExcel(Grid1, Grid2, filePath)
        End If

    End Sub



    Private Sub LoadCSVIntoDataGridView(filePath As String)
        Try
            Dim dt As New Data.DataTable()

            ' Read all lines from the CSV file
            Dim lines As String() = File.ReadAllLines(filePath)

            If lines.Length > 0 Then
                ' Create headers from the first row
                Dim headers As String() = lines(0).Split(","c)
                For Each header As String In headers
                    dt.Columns.Add(header.Trim())
                Next

                ' Add data rows
                For i As Integer = 1 To lines.Length - 1
                    Dim rowData As String() = lines(i).Split(","c)
                    dt.Rows.Add(rowData)
                Next

                ' Bind DataTable to DataGridView
                Grid1.DataSource = dt
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading CSV: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadTablesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadTablesToolStripMenuItem.Click
        Dim openFileDialog As New OpenFileDialog With {
        .Filter = "XLSX Files|*.xlsx",
        .Title = "Select a Excel File"
    }

        openFileDialog.Filter = "Excel Files|*.xlsx;*.xls"
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            LoadExcelData(openFileDialog.FileName)
        End If

    End Sub

    ' Load data from an Excel file
    Private Sub LoadExcelData(filePath As String)
        Try
            xlApp = New Excel.Application()
            xlWorkbook = xlApp.Workbooks.Open(filePath)

            ' Loop through each sheet in the workbook
            For sheetIndex As Integer = 1 To xlWorkbook.Sheets.Count
                xlWorksheet = CType(xlWorkbook.Sheets(sheetIndex), Worksheet)
                xlRange = xlWorksheet.UsedRange

                ' Create a new DataTable for the DataGridView
                Dim dt As New Data.DataTable()

                ' Ensure there are columns before adding rows
                Dim columnCount As Integer = xlRange.Columns.Count

                ' Add columns based on the first row in the sheet
                For col As Integer = 1 To columnCount
                    Dim columnName As String = If(xlRange.Cells(1, col)?.Value IsNot Nothing, xlRange.Cells(1, col).Value.ToString(), "Column" & col)
                    If Not dt.Columns.Contains(columnName) Then
                        dt.Columns.Add(columnName)
                    Else
                        dt.Columns.Add(columnName & "_" & col) ' Prevent duplicate column names
                    End If
                Next

                ' Add rows from the sheet (including duplicates)
                For row As Integer = 2 To xlRange.Rows.Count
                    Dim newRow As DataRow = dt.NewRow()

                    For col As Integer = 1 To columnCount
                        Dim cellValue As Object = xlRange.Cells(row, col)?.Value
                        newRow(col - 1) = If(cellValue IsNot Nothing, cellValue.ToString(), "")
                    Next

                    dt.Rows.Add(newRow)
                Next

                ' Bind the data to the DataGridView
                Grid1.DataSource = dt
            Next

        Catch ex As Exception
            MessageBox.Show("Error loading Excel file: " & ex.Message)
        Finally
            ' Release Excel objects
            If xlRange IsNot Nothing Then Marshal.ReleaseComObject(xlRange)
            If xlWorksheet IsNot Nothing Then Marshal.ReleaseComObject(xlWorksheet)
            If xlWorkbook IsNot Nothing Then
                xlWorkbook.Close(False)
                Marshal.ReleaseComObject(xlWorkbook)
            End If
            If xlApp IsNot Nothing Then
                xlApp.Quit()
                Marshal.ReleaseComObject(xlApp)
            End If
        End Try
    End Sub


    ' Button click event to select a file and load data

End Class