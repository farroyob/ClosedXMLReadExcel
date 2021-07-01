using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Win48_LeerExcel
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private Boolean ValidateRead()
        {
            String extension;

            if (txtDir.Text.Trim() == "")
            {
                MessageBox.Show("Seleccione el archivo antes de continuar", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return false;
            }

            if (!System.IO.File.Exists(txtDir.Text.Trim()))
            {
                MessageBox.Show("Archivo NO Existe", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return false;
            }

            extension = System.IO.Path.GetExtension(txtDir.Text.Trim());

            if (!extension.ToUpper().Contains("XLS"))
            {
                MessageBox.Show("Archivo NO es válido", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return false;
            }

            return true;
        }

        private void Read()
        {
            String fileName;
            IXLRows nonEmptyDataRows;
            DataTable dataExcel;
            Int32 currentRow;
            Int32 lineCell;
            String headerCell;
            List<String> headerList;
            Char charASCII;
            Type vloTipoDato;

            if (!ValidateRead())
            {
                return;
            }

            headerList = new List<String>();
            fileName = txtDir.Text.Trim();
            dataExcel = new DataTable("Datos");
            currentRow = 1;
            lineCell = 1;

            using (var excelWorkbook = new XLWorkbook(fileName))
            {
                nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                foreach (IXLRow dataRow in nonEmptyDataRows)
                {
                    if (currentRow == 1)
                    {
                        do
                        {
                            try
                            {
                                headerCell = (String)dataRow.Cell(lineCell).Value;

                                headerCell = headerCell.ToLower();

                                for (Int32 currentRowASCII = 0; currentRowASCII < 255; currentRowASCII++)
                                {
                                    if (currentRowASCII < 97 || currentRowASCII > 122)
                                    {
                                        charASCII = (char)currentRowASCII;

                                        headerCell = headerCell.Replace(charASCII.ToString(), "");
                                    }
                                }

                                headerCell = headerCell.ToUpper();

                                if (headerCell != "")
                                {
                                    headerList.Add(headerCell);
                                }
                            }
                            catch { headerCell = String.Empty; }

                            lineCell += 1;
                        }
                        while (headerCell != "");
                    }


                    try
                    {
                        if (currentRow >= 2)
                        {
                            if (dataExcel.Columns.Count == 0)
                            {
                                lineCell = 1;

                                foreach (String vloTitulo in headerList)
                                {
                                    vloTipoDato = dataRow.Cell(lineCell).Value.GetType();

                                    dataExcel.Columns.Add(vloTitulo, vloTipoDato);

                                    lineCell += 1;
                                }
                            }

                            lineCell = 1;

                            dataExcel.Rows.Add();

                            foreach (String headerCellColumna in headerList)
                            {
                                dataExcel.Rows[dataExcel.Rows.Count - 1][headerCellColumna] = dataRow.Cell(lineCell).Value;

                                lineCell += 1;
                            }
                        }
                    }
                    catch { headerCell = String.Empty; }

                    currentRow += 1;
                }
            }

            dtgData.DataSource = dataExcel.Copy();

            for (int i = 0; i <= dtgData.Columns.Count - 1; i++)
            {
                dtgData.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            Read();
        }

        private void lblFindFile_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.xls)|*.xls|All files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    txtDir.Text = openFileDialog.FileName;
                }
            }
        }
    }
}
