using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser.Opencut
{
    /// <summary>
    /// Used to read in data from the opencut bond migration data spreadsheet
    /// </summary>
    public class BondMigrationDataExtractor : BaseParser, IDisposable
    {

        public BondMigrationDataExtractor(string filePath): base(filePath)
        {
            base.SwitchWorksheet(1);
        }

        /// <summary>
        /// Gets a list of migration data rows from the spreadsheet
        /// </summary>
        public List<BondMigrationRow> BondData
        {
            get
            {
                int rowCount = base.ReadColumn("A", 2).Count();

                List<BondMigrationRow> data = new List<BondMigrationRow>();

                string[] columns = new string[]{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"};

                for (int i = 2; i <= rowCount + 1; i++)
                {
                    List<String> rowValues = base.ReadRow(i, columns);
                    BondMigrationRow row = new BondMigrationRow(rowValues);
                    data.Add(row);
                    Console.WriteLine(string.Format("Loading Bond Record {0} of {1}: {2} {3} {4}", i, rowCount, row.OpencutNumber, row.DEQIdentifier, row.BondNumber));
                }
                
                return data;
            }
        }


    }

    /// <summary>
    /// Encapsulates data from a single row in the bond migration data spreadsheet
    /// </summary>
    public class BondMigrationRow
    {
        public BondMigrationRow(List<string> Values)
        {
            OpencutNumber = Convert.ToInt32(Values[0]);
            DEQIdentifier = Values[1];
            BondNumber = Values[2];
            BondForm = Values[3];
            if (!string.IsNullOrEmpty(Values[4])) { CreatedDate = Convert.ToDateTime(Values[4]); }
            BondAmount = Convert.ToDecimal(Values[5]);
            if (!string.IsNullOrEmpty(Values[6])) { DEQRequiredAmount = Convert.ToDecimal(Values[6]); }
            BondType = Values[7];
            BondStatus = Values[8];
            if (!string.IsNullOrEmpty(Values[9])) { ExpirationDate = Convert.ToDateTime(Values[9]); }
            Company = Values[10];
            Address = Values[11];
            City = Values[12];
            State = Values[13];
            ZipCode = Values[14];
            if (!string.IsNullOrEmpty(Values[15])) { NaicNo = Convert.ToInt32(Values[15]); }
        }
        
        public int OpencutNumber { get; set; }
        public string DEQIdentifier { get; set; }
        public string BondNumber { get; set; }
        public string BondForm { get; set; }
        public DateTime? CreatedDate { get; set; }
        public decimal BondAmount { get; set; }
        public decimal? DEQRequiredAmount { get; set; }
        public string BondType { get; set; }
        public string BondStatus { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string Company { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public int NaicNo { get; set; }
    }
}
