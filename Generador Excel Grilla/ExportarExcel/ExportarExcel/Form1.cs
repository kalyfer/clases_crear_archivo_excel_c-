using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExportarExcel
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void BtnAgregarFila_Click(object sender, EventArgs e)
        {
            //aca recibiria los datos provenientes de la compuerta com
            byte[] fila = { 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41 };

            //crear una fila con los datos y agregarlos al data grid view
            dgvRegistros.Rows.Add(fila[0], fila[1], fila[2], fila[3], fila[4], fila[5], fila[6], fila[7], fila[8], fila[9], fila[10]);


        }

        private void BtnExportarExcel_Click(object sender, EventArgs e)
        {
            // es importante definir el nombre de las columnas que se encuentran en el data grid view
            DataTable dt = new DataTable();
            dt.Columns.Add("Column1");
            dt.Columns.Add("Column2");
            dt.Columns.Add("Column3");
            dt.Columns.Add("Column4");
            dt.Columns.Add("Column5");
            dt.Columns.Add("Column6");
            dt.Columns.Add("Column7");
            dt.Columns.Add("Column8");
            dt.Columns.Add("Column9");
            dt.Columns.Add("Column10");
            dt.Columns.Add("Column11");

            foreach (DataGridViewRow rowGrid in dgvRegistros.Rows)
            {
                //recorrer todos los registro de la grilla
                DataRow row = dt.NewRow();
                row["column1"] = rowGrid.Cells[0].Value;
                row["column2"] = rowGrid.Cells[1].Value;
                row["column3"] = rowGrid.Cells[2].Value;
                row["column4"] = rowGrid.Cells[3].Value;
                row["column5"] = rowGrid.Cells[4].Value;
                row["column6"] = rowGrid.Cells[5].Value;
                row["column7"] = rowGrid.Cells[6].Value;
                row["column8"] = rowGrid.Cells[7].Value;
                row["column9"] = rowGrid.Cells[8].Value;
                row["column10"] = rowGrid.Cells[9].Value;
                row["column11"] = rowGrid.Cells[10].Value;

                dt.Rows.Add(row);
            }

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Excel.Workbook LibroExcel;
            Excel.Worksheet HojaExcel;
            object misValue = System.Reflection.Missing.Value;

            LibroExcel = xlApp.Workbooks.Add(misValue);
            HojaExcel = (Excel.Worksheet)LibroExcel.Worksheets.get_Item(1);

            //crear cabecera excel
            HojaExcel.Cells[1, 1] = "Column1";
            HojaExcel.Cells[1, 2] = "Column2";
            HojaExcel.Cells[1, 3] = "Column3";
            HojaExcel.Cells[1, 4] = "Column4";
            HojaExcel.Cells[1, 5] = "Column5";
            HojaExcel.Cells[1, 6] = "Column6";
            HojaExcel.Cells[1, 7] = "Column7";
            HojaExcel.Cells[1, 8] = "Column8";
            HojaExcel.Cells[1, 9] = "Column9";
            HojaExcel.Cells[1, 10] = "Column10";
            HojaExcel.Cells[1, 11] = "Column11";

            //agregar toda la data del data table al archivo excel
            int aux = 2;
            foreach (DataRow dataRow in dt.Rows)
            {
                HojaExcel.Cells[aux, 1] = dataRow[0].ToString();
                HojaExcel.Cells[aux, 2] = dataRow[1].ToString();
                HojaExcel.Cells[aux, 3] = dataRow[2].ToString();
                HojaExcel.Cells[aux, 4] = dataRow[3].ToString();
                HojaExcel.Cells[aux, 5] = dataRow[4].ToString();
                HojaExcel.Cells[aux, 6] = dataRow[5].ToString();
                HojaExcel.Cells[aux, 7] = dataRow[6].ToString();
                HojaExcel.Cells[aux, 8] = dataRow[7].ToString();
                HojaExcel.Cells[aux, 9] = dataRow[8].ToString();
                HojaExcel.Cells[aux, 10] = dataRow[9].ToString();
                HojaExcel.Cells[aux, 11] = dataRow[10].ToString();
                aux++;
            }

            //cerrar el libro excel
            LibroExcel.SaveAs("C:\\prueba\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal);
            LibroExcel.Close(true, misValue, misValue);
            xlApp.Quit();

            //limpiar la y eleminar datos en memoria
            Marshal.ReleaseComObject(HojaExcel);
            Marshal.ReleaseComObject(LibroExcel);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Proceso creacion excel realizado exitosamente XD!");
        }
    }
}
