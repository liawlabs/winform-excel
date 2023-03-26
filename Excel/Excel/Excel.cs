using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//https://www.c-sharpcorner.com/UploadFile/ankurmee/import-data-from-excel-to-datagridview-in-C-Sharp/https://www.c-sharpcorner.com/UploadFile/ankurmee/import-data-from-excel-to-datagridview-in-C-Sharp/
//https://www.freecodespot.com/blog/csharp-import-excel/
//https://www.youtube.com/watch?v=LDq4C_wF0fs
//https://www.c-sharpcorner.com/UploadFile/hrojasara/export-datagridview-to-excel-in-C-Sharp/

namespace Excel
{
    public partial class Excel : Form
    {
        string filename;

        Microsoft.Office.Interop.Excel.Application xlsApp;
        Microsoft.Office.Interop.Excel.Workbook xlsWorkbook;
        Microsoft.Office.Interop.Excel.Worksheet xlsWorkSheet;
        Microsoft.Office.Interop.Excel.Range xlsRange;

        public Excel()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            openFD.Filter = "Excel Office | *.xls; *.xlsx";
            var result =  openFD.ShowDialog();
            if (result == DialogResult.OK)
            {
                filename = openFD.FileName;
                try
                {
                    xlsApp = new Microsoft.Office.Interop.Excel.Application();
                    xlsWorkbook = xlsApp.Workbooks.Open(filename);
                    xlsWorkSheet = xlsWorkbook.Worksheets["Users"];
                    xlsRange = xlsWorkSheet.UsedRange;

                    for (int row = 2; row <= xlsRange.Rows.Count; row++)
                    {
                        dgView.Rows.Add(xlsRange.Cells[row,1].Text, xlsRange.Cells[row, 2].Text, xlsRange.Cells[row, 3].Text, xlsRange.Cells[row, 4].Text);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(xlsRange);
                    Marshal.ReleaseComObject(xlsWorkSheet);
                    //quit apps
                    xlsWorkbook.Close();
                    Marshal.ReleaseComObject(xlsWorkbook);
                    xlsApp.Quit();
                    Marshal.ReleaseComObject(xlsApp);

                    btnSave.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    MessageBox.Show("Check Office x86 or x64.  Must be compatible with App");
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            xlsApp = new Microsoft.Office.Interop.Excel.Application();
            xlsWorkbook = xlsApp.Workbooks.Open(filename);
            xlsWorkSheet = xlsWorkbook.Worksheets["Users"];
            xlsRange = xlsWorkSheet.UsedRange;

            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dgView.Rows.Count; i++)
            {
                for (int j = 0; j < dgView.Columns.Count; j++)
                {
                    xlsRange.Cells[i + 2, j + 1] = dgView.Rows[i].Cells[j].Value.ToString();
                }
            }
            // save the application  
            //xlsWorkbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlsWorkbook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, System.Reflection.Missing.Value, Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
            xlsWorkbook.Saved = true;

            MessageBox.Show("File Updated");

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlsRange);
            Marshal.ReleaseComObject(xlsWorkSheet);
            //quit apps
            xlsWorkbook.Close();
            Marshal.ReleaseComObject(xlsWorkbook);
            xlsApp.Quit();
            Marshal.ReleaseComObject(xlsApp);
        }
    }
}
