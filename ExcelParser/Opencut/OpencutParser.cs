using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser.Opencut
{


    public class OpencutParser : BaseParser, IDisposable
    {

        /// <summary>
        /// Opens a new instance of the Opencut Parser class
        /// </summary>
        /// <param name="ExcelFilePath"></param>
        public OpencutParser(string ExcelFilePath) : base(ExcelFilePath) { }


        #region DataSet Methods


        /// <summary>
        /// Gets a dataset representing this spreadsheet
        /// </summary>
        /// <returns></returns>
        public DataTable GetDataset(int OpencutNumber)
        {
            DataTable table = GetEmptyTable("OpencutBoundaryCoordinates");
            AddPermitData(OpencutNumber, ref table);
            AddNonBondedData(OpencutNumber, ref table);
            AddReleaseRequestData(OpencutNumber, ref table);
            return table;
        }


        DataTable GetEmptyTable(string tableName)
        {
            DataTable table = new DataTable(tableName);
            table.Columns.Add("OpencutNumber", typeof(int));
            table.Columns.Add("Type", typeof(string));
            table.Columns.Add("SubType", typeof(string));
            table.Columns.Add("MapId", typeof(string));
            table.Columns.Add("Latitude", typeof(double));
            table.Columns.Add("Longitude", typeof(double));
            //table.Columns.Add("CollectedBy", typeof(string));
            //table.Columns.Add("CollectedDate", typeof(DateTime));
            table.Columns.Add("Description", typeof(string));
            return table;

        }

        void AddPermitData(int OpencutNumber, ref DataTable table)
        {
            foreach (GeoCoordinate coord in PermitBoundryCoordinates)
            {
                table.Rows.Add(OpencutNumber, "Boundary", "Operator Proposed Permit", coord.MapId, coord.Latitude, coord.Longitude, coord.Description);
            }
        }

        void AddNonBondedData(int OpencutNumber, ref DataTable table)
        {
            foreach (GeoCoordinate coord in NonBondedBoundryCoordinates)
            {
                table.Rows.Add(OpencutNumber, "Boundary", "Operator Proposed Non-Bonded", coord.MapId, coord.Latitude, coord.Longitude, coord.Description);
            }
        }

        void AddReleaseRequestData(int OpencutNumber, ref DataTable table)
        {
            string purpose = string.Empty;
            if (ReleaseRequestPurpose == "Acreage Release")
                purpose = "Operator Proposed Release Request Phase II";
            if (ReleaseRequestPurpose == "Bond Reduction")
                purpose = "Operator Proposed Release Request Phase I";
            if (ReleaseRequestPurpose == "Acreage Release & Bond Reduction")
                throw new Exception("'Acreage Release & Bond Reduction' Releast Requests must be done manually. The coordinates have not been included in the data.");
            

            foreach (GeoCoordinate coord in ReleaseRequestBoundryCoordinates)
            {
                table.Rows.Add(OpencutNumber, "Boundary", purpose, coord.MapId, coord.Latitude, coord.Longitude, coord.Description);
            }
        }

        #endregion


        /// <summary>
        /// Gets the approximate center of site
        /// </summary>
        public GeoCoordinate CenterOfSite
        {
            get
            {
                base.SwitchWorksheet(2);
                return new GeoCoordinate(base.GetValue("B26"), base.GetValue("C26"), base.GetValue("A26"), base.GetValue("D26"));
            }
        }


        /// <summary>
        /// The operator name
        /// </summary>
        public string OperatorName
        {
            get
            {
                base.SwitchWorksheet(2);
                return base.GetValue("C19");
            }
        }


        /// <summary>
        /// The site name
        /// </summary>
        public string SiteName
        {
            get
            {
                base.SwitchWorksheet(2);
                return base.GetValue("C21");
            }
        }


        /// <summary>
        /// The permit number
        /// </summary>
        public string PermitNumber
        {
            get
            {
                base.SwitchWorksheet(2);
                return base.GetValue("C23");
            }
        }


        /// <summary>
        /// The date entered by operator
        /// </summary>
        public string Date
        {
            get
            {
                base.SwitchWorksheet(2);
                return base.GetValue("E23");
            }
        }


        /// <summary>
        /// The purpose of the permit boundary coordinate form
        /// </summary>
        public string PermitPurpose
        {
            get
            {
                base.SwitchWorksheet(2);
                return base.GetValue("D4");
            }
        }

        /// <summary>
        /// The purpose of the non bonded boundary coordinate form
        /// </summary>
        public string NonBondedPurpose
        {
            get
            {
                base.SwitchWorksheet(3);
                return base.GetValue("D3");
            }
        }

        /// <summary>
        /// The purpose of the Release Request boundary coordinate form
        /// </summary>
        public string ReleaseRequestPurpose
        {
            get
            {
                base.SwitchWorksheet(4);
                return base.GetValue("D3");
            }
        }


        /// <summary>
        /// returns the coordinates from the Permit Boundry worksheet
        /// </summary>
        public List<GeoCoordinate> PermitBoundryCoordinates
        {
            get
            {
                List<GeoCoordinate> coords = new List<GeoCoordinate>();

                base.SwitchWorksheet(2);

                List<string> mapIds = base.ReadColumn("A", 27);
                List<string> lats = base.ReadColumn("B", 27);
                List<string> lngs = base.ReadColumn("C", 27);
                List<string> descs = base.ReadColumn("D", 27, lats.Count);


                for (int i = 0; i < lats.Count; i++)
                {
                    double latitude = Convert.ToDouble(lats[i]);
                    double longitude = Convert.ToDouble(lngs[i]);
                    GeoCoordinate coord = new GeoCoordinate(latitude, longitude, mapIds[i], descs[i]);
                    coords.Add(coord);
                }

                return coords;
            }
        }


        /// <summary>
        /// returns the coordinates from the Non-Bonded boundry worksheet
        /// </summary>
        public List<GeoCoordinate> NonBondedBoundryCoordinates
        {
            get
            {
                List<GeoCoordinate> coords = new List<GeoCoordinate>();

                base.SwitchWorksheet(3);

                List<string> ids = base.ReadColumn("A", 22);
                List<string> lats = base.ReadColumn("B", 22);
                List<string> lngs = base.ReadColumn("C", 22);
                List<string> descs = base.ReadColumn("D", 22, lats.Count);

                for (int i = 0; i < lats.Count; i++)
                {
                    double latitude = Convert.ToDouble(lats[i]);
                    double longitude = Convert.ToDouble(lngs[i]);
                    GeoCoordinate coord = new GeoCoordinate(latitude, longitude, ids[i], descs[i]);
                    coords.Add(coord);
                }

                return coords;
            }
        }


        /// <summary>
        /// returns the coordinates from the Release Request Boundry worksheet
        /// </summary>
        public List<GeoCoordinate> ReleaseRequestBoundryCoordinates
        {
            get
            {
                List<GeoCoordinate> coords = new List<GeoCoordinate>();

                base.SwitchWorksheet(4);

                List<string> ids = base.ReadColumn("A", 20);
                List<string> lats = base.ReadColumn("B", 20);
                List<string> lngs = base.ReadColumn("C", 20);
                List<string> descs = base.ReadColumn("D", 20, lats.Count);

                for (int i = 0; i < lats.Count; i++)
                {
                    double latitude = Convert.ToDouble(lats[i]);
                    double longitude = Convert.ToDouble(lngs[i]);
                    GeoCoordinate coord = new GeoCoordinate(latitude, longitude, ids[i], descs[i]);
                    coords.Add(coord);
                }

                return coords;
            }
        }

    }


}
