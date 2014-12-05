using ExcelParser.Opencut;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpencutBoundaryCoordinates
{
    public partial class Form1 : Form
    {

        string _filePath = string.Empty;

        #region Form Event Handlers

        public Form1()
        {
            InitializeComponent();
            openFileDialog1.FileOk += openFileDialog1_FileOk;
        }

        /// <summary>
        /// opens a dialog for selecting a spreadsheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
        }

        /// <summary>
        /// updates the file path varialbe when a spreadsheet file is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            _filePath = openFileDialog1.FileName;
            txtSpreadsheetPath.Text = _filePath;
        }
        
        /// <summary>
        /// validates the form and displays the data if valid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLoadData_Click(object sender, EventArgs e)
        {
            if (IsFormValid())
                DisplayData();
        }

        /// <summary>
        /// deletes selected rows from the data grid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeleteRows_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection rows = dataGridView1.SelectedRows;
            foreach (DataGridViewRow row in rows)
                dataGridView1.Rows.Remove(row);
            
        }

        /// <summary>
        /// calls the export function
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportData_Click(object sender, EventArgs e)
        {
            ExportData();
        }

        #endregion


        /// <summary>
        /// Ensures a file is selected and an opencut number is provided
        /// </summary>
        /// <returns></returns>
        bool IsFormValid()
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                MessageBox.Show("You must select a spreadsheet", "Insufficient Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrEmpty(txtOpencutNumber.Text))
            {
                MessageBox.Show("You must enter an Opencut Number", "Insufficient Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            int myNum = 0;
            if (!Int32.TryParse(txtOpencutNumber.Text, out myNum))
            {
                MessageBox.Show("The Opencut Number must be an integer", "Incorrect Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }


        /// <summary>
        /// displays data from the spreadsheet
        /// </summary>
        void DisplayData()
        {
            int OpencutNumber = Convert.ToInt32(txtOpencutNumber.Text);
            OpencutParser parser = new OpencutParser(_filePath);
            DataTable data = parser.GetDataset(OpencutNumber);
            dataGridView1.DataSource = data;
            dataGridView1.Show();
        }


        /// <summary>
        /// Exports the data for the boundary coordinate drawing program
        /// </summary>
        void ExportData()
        {
            string csv = GetCSVData();

            Clipboard.SetText(csv);

            MessageBox.Show("The data has been copied to the clipboard");
        }


        /// <summary>
        /// Formats the data into CSV
        /// </summary>
        /// <returns></returns>
        string GetCSVData()
        {
            StringBuilder sb = new StringBuilder();
            DateTime exportDate = DateTime.Now;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                List<string> values = new List<string>();

                // add data from the cells
                for (int i = 0; i < 6; i++)
                {
                    DataGridViewCell cell = row.Cells[i];
                    if (cell.Value == null)
                        break;
                    values.Add(cell.Value.ToString());
                }

                // if the cells contain data append the collected by and collected date values
                if (values.Count > 0)
                {
                    values.Add("Operator");
                    values.Add(exportDate.ToShortDateString());
                    sb.AppendLine(string.Join(",", values.ToArray()));
                }
            }

            return sb.ToString().Trim();
        }

        

    }
}
