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
    public class InactiveSiteExtractor : BaseParser, IDisposable
    {

        public InactiveSiteExtractor(string filePath) : base(filePath)
        {
            base.SwitchWorksheet(1);
        }

        /// <summary>
        /// Gets the other sites data
        /// </summary>
        public List<SiteRow> Sites
        {
            get
            {
                base.SwitchWorksheet(1);

                int rowCount = base.ReadColumn("K", 2).Count();
                List<SiteRow> data = new List<SiteRow>();

                string[] columns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"};

                for (int i = 2; i <= rowCount; i++)
                {
                    List<String> rowValues = base.ReadRow(i, columns);
                    SiteRow row = new SiteRow(rowValues);
                    data.Add(row);
                    Console.WriteLine(string.Format("Loading Site Record {0} of {1}: {2} {3} {4}", i, rowCount, row.CompanyId, row.CompanyName, row.SiteName));
                }

                return data;
            }
        }


        /// <summary>
        /// Gets the other sites data
        /// </summary>
        public List<OtherSiteRow> OtherSites
        {
            get
            {
                base.SwitchWorksheet(2);

                int rowCount = base.ReadColumn("K", 2).Count();
                List<OtherSiteRow> data = new List<OtherSiteRow>();

                string[] columns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S" };

                for (int i = 2; i <= rowCount + 1; i++)
                {
                    List<String> rowValues = base.ReadRow(i, columns);
                    OtherSiteRow row = new OtherSiteRow(rowValues);
                    data.Add(row);
                    Console.WriteLine(string.Format("Loading Other Site Record {0} of {1}: {2} {3} {4}", i, rowCount, row.CompanyId, row.CompanyName, row.SiteName));
                }

                return data;
            }
        }

        


    }

    public class SiteRow
    {
        public SiteRow(List<string> Values)
        {
            if (!string.IsNullOrEmpty(Values[0])) { CompanyId = Convert.ToInt32(Values[0]); }
            PermitNumber = Values[1];
            CompanyName = Values[2];
            OperatorType = Values[3];
            Phone = Values[4];
            Address1 = Values[5];
            Address2 = Values[6];
            City = Values[7];
            Region = Values[8];
            ZipCode = Values[9];
            // Country USA is value 10
            SiteId = Convert.ToInt32(Values[11]);
            SiteName = Values[12];
            Status = Values[13];
            County = Values[14];
            if (!string.IsNullOrEmpty(Values[15])) { Latitude = Convert.ToDecimal(Values[15]); }
            if (!string.IsNullOrEmpty(Values[16])) { Longitude = Convert.ToDecimal(Values[16]); }
        }

        public int? CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string OperatorType { get; set; }
        public string Phone { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string ZipCode { get; set; }
        public int SiteId { get; set; }
        public string PermitNumber { get; set; }
        public string SiteName { get; set; }
        public string Status { get; set; }
        public string County { get; set; }
        public decimal? Latitude { get; set; }
        public decimal? Longitude { get; set; }

    }


    /// <summary>
    /// A row of data from the other sites tab
    /// </summary>
    public class OtherSiteRow
    {
        public OtherSiteRow(List<string> Values)
        {
            if (!string.IsNullOrEmpty(Values[0])) { CompanyId = Convert.ToInt32(Values[0]); }
            CompanyName = Values[1];
            OperatorType = Values[2];
            Phone = Values[3];
            Address1 = Values[4];
            Address2 = Values[5];
            City = Values[6];
            Region = Values[7];
            ZipCode  = Values[8];
            SFID = Convert.ToInt32(Values[10]);
            SFPermit = Values[11];
            if (!string.IsNullOrEmpty(Values[12])) { PermitNumber = Convert.ToInt32(Values[12]); }
            SiteName = Values[13];            
            Status = Values[14];
            County = Values[15];
            if (!string.IsNullOrEmpty(Values[16])) { Latitude = Convert.ToDecimal(Values[16]); }
            if (!string.IsNullOrEmpty(Values[17])) { Longitude = Convert.ToDecimal(Values[17]); }
            AppType = Values[18];
        }

        public int? CompanyId { get; set; } 
        public string CompanyName { get; set; }
        public string OperatorType { get; set; }
        public string Phone { get; set; } 
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; } 
        public string Region { get; set; } 
        public string ZipCode { get; set; }
        public int SFID { get; set; } 
        public string SFPermit { get; set; } 
        public int? PermitNumber { get; set; }
        public string SiteName { get; set; } 
        public string Status { get; set; } 
        public string County { get; set; } 
        public decimal? Latitude { get; set; }
        public decimal? Longitude { get; set; }
        public string AppType { get; set; } 
        


    }
}
