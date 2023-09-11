using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
namespace APP
{
    public partial class Form1 : Form
    {
        conexion conexion = new conexion();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1º_Load(object sender, EventArgs e)
        {
           
        }

        public void llenarCampos()
        {
            string consulta = "select * from Empleados";
            SqlDataAdapter adaptador = new SqlDataAdapter(consulta, conexion.cnn);
            System.Data.DataTable dt = new System.Data.DataTable();
            adaptador.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        public void ExportarDatos(DataGridView datalistado)
        {
            Microsoft.Office.Interop.Excel.Application exportarexcel = new Microsoft.Office.Interop.Excel.Application();
            exportarexcel.Application.Workbooks.Add(true);

            int indicecolumna = 0;

            foreach (DataGridViewColumn columna in datalistado.Columns)
            {
                indicecolumna++;
                exportarexcel.Cells[1, indicecolumna] = columna.Name;
            }
            
            int indicefila = 0;
            foreach (DataGridViewRow fila in datalistado.Rows )
            {
                indicefila++;
                indicecolumna = 0;
                foreach (DataGridViewColumn columna in datalistado.Columns)
                {
                    indicecolumna++;
                    exportarexcel.Cells[indicefila + 1, indicecolumna] = fila.Cells[columna.Name].Value;
                }
            }

            exportarexcel.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ExportarDatos(dataGridView1);
            ExportarDataGridViewAExcel(dataGridView1);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            llenarCampos();
        }


        private void ExportarDataGridViewAExcel(DataGridView dgv)
        {
            try
            {
                // Crear una instancia de Excel
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;

                // Crear un nuevo libro de Excel
                Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                Worksheet worksheet = workbook.Sheets[1];

                // Copiar los datos desde el DataGridView a Excel
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 1, j + 1] = dgv.Rows[i].Cells[j].Value;
                    }
                }

                // Guardar el archivo de Excel (opcional)
                // workbook.SaveAs("ruta_del_archivo.xlsx");

                // Liberar recursos
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                MessageBox.Show("Datos exportados correctamente a Excel", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar los datos a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
