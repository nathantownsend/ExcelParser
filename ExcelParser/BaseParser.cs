using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    public class BaseParser: IDisposable
    {
        //source: http://forum.codecall.net/topic/71788-reading-excel-files-in-c/
        private Microsoft.Office.Interop.Excel.Application appExcel;
        private Workbook newWorkbook = null;
        private _Worksheet objsheet = null;


        /// <summary>
        /// Creates an instance of the base parser class for an Excel file
        /// </summary>
        /// <param name="excelFilePath"></param>
        public BaseParser(string excelFilePath)
        {
            ExcelFilePath = excelFilePath;
            excel_init(ExcelFilePath);
        }


        /// <summary>
        /// The path to the Excel file
        /// </summary>
        public string ExcelFilePath { get; set; }


        /// <summary>
        /// The index of the worksheet
        /// </summary>
        /// <param name="index"></param>
        public void SwitchWorksheet(int index)
        {
            objsheet = (_Worksheet)appExcel.ActiveWorkbook.Worksheets[index];
        }


        /// <summary>
        /// Returns values from the column starting at the specified row and continues until an empty row is encountered
        /// </summary>
        /// <param name="StartingRow"></param>
        /// <returns></returns>
        public List<string> ReadColumn(string Column, int StartingRow)
        {
            List<string> values = new List<string>();

            string value = GetValue(Column, StartingRow);
            while (!string.IsNullOrEmpty(value))
            {
                values.Add(value);
                StartingRow += 1;
                value = GetValue(Column, StartingRow);
            }

            return values;
        
        }


        /// <summary>
        /// Returns values from the column starting a the specified row and continuing for the number of rows specified
        /// </summary>
        /// <param name="Column"></param>
        /// <param name="StartingRow"></param>
        /// <param name="NumberOfRows"></param>
        /// <returns></returns>
        public List<string> ReadColumn(string Column, int StartingRow, int NumberOfRows)
        {
            List<string> values = new List<string>();

            string value = GetValue(Column, StartingRow);
            for(int row = StartingRow; row <= StartingRow + NumberOfRows; row++)
            {
                value = GetValue(Column, row);
                values.Add(value);
            }

            return values;
        }


        /// <summary>
        /// Returns values for each column from the row specified
        /// </summary>
        /// <param name="Row"></param>
        /// <param name="ColumnsToRead"></param>
        /// <returns></returns>
        public List<string> ReadRow(int Row, string[] Columns) 
        {
            List<string> values = new List<string>();

            foreach (string col in Columns)
                values.Add(GetValue(col, Row));

            return values;
        }
        

        #region Internal Ops


        // Initializes a new instance of excel interop, opens the document, and loads the active sheet
        void excel_init(String path)
        {
            appExcel = new Microsoft.Office.Interop.Excel.Application();

            if (System.IO.File.Exists(path))
            {
                newWorkbook = appExcel.Workbooks.Open(path, true, true);
                objsheet = (_Worksheet)appExcel.ActiveWorkbook.ActiveSheet;
            }
            else
            {
                Console.WriteLine("Unable to open file!");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
                appExcel = null;
            }

        }


        /// <summary>
        /// Returns the value of a cell
        /// </summary>
        /// <param name="Column"></param>
        /// <param name="Row"></param>
        /// <returns></returns>
        public string GetValue(string Column, int Row)
        {
            string cell = string.Format("{0}{1}", Column, Row);
            return GetValue(cell);
        }


        /// <summary>
        /// Returns the value of a cell
        /// </summary>
        /// <param name="cellname"></param>
        /// <returns></returns>
        public string GetValue(string cellname)
        {
            string value = string.Empty;
            try 
            { 
                value = objsheet.get_Range(cellname).get_Value().ToString(); 
            }
            catch { }

            return value;
        }


        //Method to close excel connection
        void excel_close()
        {
            if (appExcel != null)
            {
                try
                {
                    newWorkbook.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
                    appExcel = null;
                    objsheet = null;
                }
                catch (Exception ex)
                {
                    appExcel = null;
                    Console.WriteLine("Unable to release the Object " + ex.ToString());
                }
                finally
                {
                    GC.Collect();
                }
            }
        }
        

        /// <summary>
        /// Releases resources
        /// </summary>
        public void Dispose()
        {
            excel_close();
        }

        #endregion
    }
}
