Imports System.IO
Imports System.Windows.Forms
Imports ClosedXML.Excel
Imports Microsoft.Office.Interop.Excel
'Imports System.Data


Public Class Utility

    Public Sub ProExportarExcel(ByVal dt As System.Data.DataTable, ByVal stTitulo As String, Optional InitialDirectory As String = Nothing)
        Dim fileName As String

        Try

            Dim saveFileDialog As New SaveFileDialog()

            If InitialDirectory <> Nothing Then
                saveFileDialog.InitialDirectory = InitialDirectory
            End If

            saveFileDialog.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            saveFileDialog.Title = "Guardar " & stTitulo
            saveFileDialog.FileName = stTitulo ' + " (" + DateTime.Now.ToString("yyyy-MM-dd") + ")"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                fileName = saveFileDialog.FileName

                Using wb As New XLWorkbook()
                    wb.Worksheets.Add(CType(dt, Data.DataTable), stTitulo)
                    ' Adjust widths of Columns.
                    wb.Worksheet(1).Columns().AdjustToContents()
                    ' Save the Excel file.
                    wb.SaveAs(fileName)
                End Using

                'Convertir el archivo a formateo .xls
                File.Move(fileName, Path.ChangeExtension(fileName, ".xls"))

                MessageBox.Show(stTitulo & " Compilado Guardado Existosamente")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub ProExportarExcelSimple(ByVal dt As System.Data.DataTable, ByVal stTitulo As String, Optional InitialDirectory As String = Nothing)
        ' Create a SaveFileDialog
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xls" 'This Filter define archive extension
        saveFileDialog.Title = "Guardar " & stTitulo
        saveFileDialog.FileName = stTitulo


        'Archive with route
        Dim fileName As String

        If saveFileDialog.ShowDialog() = DialogResult.OK Then

            fileName = saveFileDialog.FileName

            Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
            Dim excelWorkBook As Workbook = excelApp.Workbooks.Add(Type.Missing)
            Dim excelWorkSheet As Worksheet = CType(excelWorkBook.Sheets(1), Worksheet)

            Try
                ' Add column headers
                For i As Integer = 1 To dt.Columns.Count
                    'Hacemos algunos ajustes al excel para que las referencias quede en la columna K

                    If dt.Columns(i - 1).ColumnName.Contains("Column") Then
                        excelWorkSheet.Cells(1, i) = " "
                    Else
                        excelWorkSheet.Cells(1, i) = dt.Columns(i - 1).ColumnName
                    End If

                Next

                ' Add rows
                For i As Integer = 0 To dt.Rows.Count - 1
                    For j As Integer = 0 To dt.Columns.Count - 1
                        excelWorkSheet.Cells(i + 2, j + 1) = dt.Rows(i)(j).ToString()
                    Next
                Next

                ' Save the workbook
                excelWorkBook.SaveAs(fileName)
                excelWorkBook.Close()
                excelApp.Quit()

                MessageBox.Show(stTitulo & " Compilado Guardado Existosamente")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If
    End Sub

End Class

