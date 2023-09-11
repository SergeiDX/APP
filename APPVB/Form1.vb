Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    Private conexion As New SqlConnection("Data Source=DESKTOP-EUB75BK;Initial Catalog=Capacitacion;Integrated Security=True")
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim consulta = "select * from Empleados"
            Dim adaptador As SqlDataAdapter = New SqlDataAdapter(consulta, conexion)
            Dim dt As DataTable = New DataTable()
            adaptador.Fill(CType(dt, DataSet))
            DataGridView1.DataSource = dt
            conexion.Open()
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            conexion.Close()
        End Try
    End Sub

    Public Sub ExportarDatos(ByVal datalistado As DataGridView)
        Dim exportarexcel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        exportarexcel.Application.Workbooks.Add(True)
        Dim indicecolumna As Integer = 0

        For Each columna As DataGridViewColumn In datalistado.Columns
            indicecolumna += 1
            exportarexcel.Cells(1, indicecolumna) = columna.Name
        Next

        Dim indicefila As Integer = 0

        For Each fila As DataGridViewRow In datalistado.Rows
            indicefila += 1
            indicecolumna = 0

            For Each columna As DataGridViewColumn In datalistado.Columns
                indicecolumna += 1
                exportarexcel.Cells(indicefila + 1, indicecolumna) = fila.Cells(columna.Name).Value
            Next
        Next

        exportarexcel.Visible = True
    End Sub

    Private Sub ExportarDataGridViewAExcel(ByVal dgv As DataGridView)
        Try
            Dim excelApp = New Application()
            excelApp.Visible = True
            Dim workbook As Workbook = excelApp.Workbooks.Add(Type.Missing)
            Dim worksheet As Worksheet = workbook.Sheets(1)

            For i As Integer = 0 To dgv.Rows.Count - 1

                For j As Integer = 0 To dgv.Columns.Count - 1
                    worksheet.Cells(i + 1, j + 1) = dgv.Rows(i).Cells(j).Value
                Next
            Next

            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            MessageBox.Show("Datos exportados correctamente a Excel", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al exportar los datos a Excel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'ExportarDatos(DataGridView1)
        ExportarDataGridViewAExcel(DataGridView1)
    End Sub
End Class
